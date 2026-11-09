"""Create the #601 fixture and previews using the supplied local CLI binary."""
import argparse
import os
from pathlib import Path
import subprocess

p = argparse.ArgumentParser()
p.add_argument("cli", type=Path)
p.add_argument("output", type=Path)
args = p.parse_args()
args.output.mkdir(parents=True, exist_ok=True)
env = {**os.environ, "OFFICECLI_SKIP_UPDATE": "1", "OFFICECLI_NO_AUTO_RESIDENT": "1"}
deck = args.output.resolve() / "flips.pptx"

def cli(*command):
    result = subprocess.run([str(args.cli.resolve()), *map(str, command)], env=env,
                            capture_output=True, encoding="utf-8", errors="replace")
    if result.returncode:
        raise RuntimeError(f"{command}: {result.stdout}\n{result.stderr}")
    return result.stdout

cli("create", deck)
cli("add", deck, "/", "--type", "slide", "--prop", "layout=blank")
cases = [("Normal", False, False, 0), ("Horizontal", True, False, 0),
         ("Vertical", False, True, 0), ("Both", True, True, 0),
         ("Rotate 25", False, False, 25), ("Flip + rotate", True, False, 25)]
for index, (label, horizontal, vertical, rotation) in enumerate(cases):
    props = dict(geometry="chevron", name=label, text=f"{label}\nHello 123",
                 x=f"{2 + (index % 3) * 10}cm", y=f"{3 + (index // 3) * 8}cm",
                 width="8cm", height="3cm", fill="2E6E4A", color="FFFFFF",
                 size="22", font="Arial", align="center", valign="middle",
                 flipH=str(horizontal).lower(), flipV=str(vertical).lower(),
                 rotation=str(rotation))
    command = ["add", deck, "/slide[1]", "--type", "shape"]
    for key, value in props.items():
        command.extend(["--prop", f"{key}={value}"])
    cli(*command)
cli("validate", deck)
cli("view", deck, "html", "--page", "1", "-o", args.output.resolve() / "flips.html")
# SVG mode writes to stdout; unlike HTML it does not implement --out.
(args.output / "flips.svg").write_text(cli("view", deck, "svg", "--page", "1"), encoding="utf-8")
print(deck)
