import argparse
import shutil
import struct
import subprocess
import sys
from pathlib import Path


def decompress_vba_container(data: bytes) -> bytes:
    if not data or data[0] != 0x01:
        raise ValueError("Invalid VBA compressed container signature")

    pos = 1
    out = bytearray()

    while pos + 2 <= len(data):
        header = int.from_bytes(data[pos : pos + 2], "little")
        pos += 2

        chunk_size = (header & 0x0FFF) + 3
        compressed = (header >> 15) & 0x1

        chunk_end = min(len(data), pos + chunk_size - 2)
        chunk_data = data[pos:chunk_end]
        pos = chunk_end

        if not compressed:
            out.extend(chunk_data[:4096])
            continue

        chunk_out = bytearray()
        i = 0
        while i < len(chunk_data) and len(chunk_out) < 4096:
            flags = chunk_data[i]
            i += 1

            for bit in range(8):
                if i >= len(chunk_data) or len(chunk_out) >= 4096:
                    break

                if ((flags >> bit) & 1) == 0:
                    chunk_out.append(chunk_data[i])
                    i += 1
                    continue

                if i + 1 >= len(chunk_data):
                    i = len(chunk_data)
                    break

                token = chunk_data[i] | (chunk_data[i + 1] << 8)
                i += 2

                n = len(chunk_out)
                bit_count = max(4, (n - 1).bit_length() if n > 0 else 0)
                length = (token & ((1 << (16 - bit_count)) - 1)) + 3
                offset = (token >> (16 - bit_count)) + 1

                for _ in range(length):
                    if offset > len(chunk_out):
                        raise ValueError("Invalid copy token offset")
                    chunk_out.append(chunk_out[-offset])
                    if len(chunk_out) >= 4096:
                        break

        out.extend(chunk_out)

    return bytes(out)


def read_u16(data: bytes, pos: int):
    return struct.unpack_from("<H", data, pos)[0], pos + 2


def read_u32(data: bytes, pos: int):
    return struct.unpack_from("<I", data, pos)[0], pos + 4


def read_blob(data: bytes, pos: int):
    size, pos = read_u32(data, pos)
    return data[pos : pos + size], pos + size


def parse_dir_modules(dir_plain: bytes, text_encoding: str):
    modules = []
    i = 0
    n = len(dir_plain)

    while i + 2 <= n:
        rec_id, next_i = read_u16(dir_plain, i)
        if rec_id != 0x19:
            i += 1
            continue

        try:
            i = next_i
            name_b, i = read_blob(dir_plain, i)

            rec_id, i = read_u16(dir_plain, i)
            if rec_id != 0x47:
                continue
            _, i = read_blob(dir_plain, i)

            rec_id, i = read_u16(dir_plain, i)
            if rec_id != 0x1A:
                continue
            stream_b, i = read_blob(dir_plain, i)

            rec_id, i = read_u16(dir_plain, i)
            if rec_id != 0x32:
                continue
            _, i = read_blob(dir_plain, i)

            rec_id, i = read_u16(dir_plain, i)
            if rec_id != 0x1C:
                continue
            _, i = read_blob(dir_plain, i)

            rec_id, i = read_u16(dir_plain, i)
            if rec_id != 0x48:
                continue
            _, i = read_blob(dir_plain, i)

            rec_id, i = read_u16(dir_plain, i)
            if rec_id != 0x31:
                continue
            size, i = read_u32(dir_plain, i)
            if size != 4:
                continue
            text_offset, i = read_u32(dir_plain, i)

            rec_id, i = read_u16(dir_plain, i)
            if rec_id != 0x1E:
                continue
            size, i = read_u32(dir_plain, i)
            i += size

            rec_id, i = read_u16(dir_plain, i)
            if rec_id != 0x2C:
                continue
            size, i = read_u32(dir_plain, i)
            i += size

            rec_type, i = read_u16(dir_plain, i)
            if rec_type not in (0x21, 0x22):
                continue
            _, i = read_u32(dir_plain, i)

            rec_next, j = read_u16(dir_plain, i)
            if rec_next == 0x25:
                i = j + 4
                rec_next, j = read_u16(dir_plain, i)
            if rec_next == 0x28:
                i = j + 4
                rec_next, j = read_u16(dir_plain, i)
            if rec_next != 0x2B:
                continue
            i = j + 4

            modules.append(
                {
                    "name": name_b.decode(text_encoding, errors="replace"),
                    "stream": stream_b.decode(text_encoding, errors="replace"),
                    "text_offset": text_offset,
                    "kind": "procedural" if rec_type == 0x21 else "class",
                }
            )
        except Exception:
            i = next_i

    return modules


def parse_project_codepage(dir_plain: bytes) -> int:
    # PROJECTCODEPAGE record signature:
    # id=0x0003, size=0x00000002, then 2-byte code page.
    marker = b"\x03\x00\x02\x00\x00\x00"
    idx = dir_plain.find(marker)
    if idx >= 0 and idx + 8 <= len(dir_plain):
        return int.from_bytes(dir_plain[idx + 6 : idx + 8], "little")
    return 1252


def codepage_to_python_encoding(codepage: int) -> str:
    return f"cp{codepage}"


def parse_project_types(project_text: str):
    modules = set()
    forms = set()
    documents = set()

    for line in project_text.splitlines():
        if line.startswith("Module="):
            modules.add(line.split("=", 1)[1].strip())
        elif line.startswith("BaseClass="):
            forms.add(line.split("=", 1)[1].strip())
        elif line.startswith("Document="):
            raw = line.split("=", 1)[1].strip()
            documents.add(raw.split("/", 1)[0].strip())

    return modules, forms, documents


def choose_extension(name: str, stream: str, kind: str, type_info):
    modules, forms, documents = type_info

    if name in modules or stream in modules:
        return ".bas"
    if name in forms or stream in forms:
        return ".frm"
    if name in documents or stream in documents:
        return ".cls"
    return ".bas" if kind == "procedural" else ".cls"


def unique_path(path: Path):
    if not path.exists():
        return path
    stem = path.stem
    suffix = path.suffix
    idx = 2
    while True:
        candidate = path.with_name(f"{stem}_{idx}{suffix}")
        if not candidate.exists():
            return candidate
        idx += 1


def run_7z_extract(input_bin: Path, output_dir: Path):
    seven_zip = shutil.which("7z")
    if not seven_zip:
        raise RuntimeError("7z was not found in PATH")

    cmd = [seven_zip, "x", str(input_bin), f"-o{output_dir}", "-y"]
    proc = subprocess.run(cmd, capture_output=True, text=True)
    if proc.returncode != 0:
        raise RuntimeError(
            "7z extraction failed:\n"
            + (proc.stdout or "")
            + ("\n" + proc.stderr if proc.stderr else "")
        )


def unpack_vba(
    input_bin: Path,
    out_dir: Path,
    tmp_dir: Path,
    keep_tmp: bool,
    output_encoding: str = "project",
):
    if tmp_dir.exists():
        shutil.rmtree(tmp_dir)
    tmp_dir.mkdir(parents=True, exist_ok=True)

    run_7z_extract(input_bin, tmp_dir)

    project_path = tmp_dir / "PROJECT"
    dir_path = tmp_dir / "VBA" / "dir"
    vba_dir = tmp_dir / "VBA"

    dir_plain = decompress_vba_container(dir_path.read_bytes())
    codepage = parse_project_codepage(dir_plain)
    project_encoding = codepage_to_python_encoding(codepage)
    effective_output_encoding = (
        project_encoding if output_encoding.strip().lower() == "project" else output_encoding
    )

    project_text = project_path.read_text(encoding=project_encoding, errors="replace")
    type_info = parse_project_types(project_text)

    modules = parse_dir_modules(dir_plain, text_encoding=project_encoding)
    if not modules:
        raise RuntimeError("No VBA modules found in dir stream")

    out_dir.mkdir(parents=True, exist_ok=True)

    written = []
    for m in modules:
        stream_path = vba_dir / m["stream"]
        if not stream_path.exists():
            continue

        raw_stream = stream_path.read_bytes()
        if m["text_offset"] >= len(raw_stream):
            continue

        source_bytes = decompress_vba_container(raw_stream[m["text_offset"] :])
        source_text = source_bytes.decode(project_encoding, errors="replace")

        ext = choose_extension(m["name"], m["stream"], m["kind"], type_info)
        target = unique_path(out_dir / f"{m['name']}{ext}")
        target.write_text(source_text, encoding=effective_output_encoding, newline="")
        written.append(target)

    if not keep_tmp:
        shutil.rmtree(tmp_dir, ignore_errors=True)

    return written, codepage, project_encoding, effective_output_encoding


def main():
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")

    parser = argparse.ArgumentParser(
        description="Extract VBA source code from vbaProject.bin"
    )
    parser.add_argument(
        "input_bin",
        nargs="?",
        default="vbaProject.bin",
        help="Path to vbaProject.bin",
    )
    parser.add_argument(
        "--out",
        default="vba_unpacked",
        help="Output directory for extracted code files",
    )
    parser.add_argument(
        "--tmp",
        default=".vba_extract_tmp",
        help="Temporary extraction directory",
    )
    parser.add_argument(
        "--keep-temp",
        action="store_true",
        help="Keep temporary extracted OLE tree",
    )
    parser.add_argument(
        "--encoding",
        default="project",
        help="Output encoding: 'project' (default) or any Python codec (e.g. utf-8, cp1251)",
    )
    args = parser.parse_args()

    input_bin = Path(args.input_bin).resolve()
    out_dir = Path(args.out).resolve()
    tmp_dir = Path(args.tmp).resolve()

    files, codepage, project_encoding, used_output_encoding = unpack_vba(
        input_bin=input_bin,
        out_dir=out_dir,
        tmp_dir=tmp_dir,
        keep_tmp=args.keep_temp,
        output_encoding=args.encoding,
    )
    print(
        f"Extracted {len(files)} modules to: {out_dir}\n"
        f"Project code page: {codepage}\n"
        f"Project encoding: {project_encoding}\n"
        f"Output encoding: {used_output_encoding}"
    )
    for f in files:
        print(f"- {f.name}")


if __name__ == "__main__":
    main()
