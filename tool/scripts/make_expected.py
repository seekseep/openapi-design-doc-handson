"""Write <name>.expected next to each fixture, using PyYAML as the reference.

The format matches modYaml.YamlDump: JSON-like, every scalar is a string
(PyYAML BaseLoader), null / empty values are "".
fixtures/openapi.yaml is a copy of the sample used in the hands-on.
"""
import sys
from pathlib import Path

import yaml

ROOT = Path(__file__).resolve().parents[2]


def dump(v) -> str:
    if v is None:
        return '""'
    if isinstance(v, str):
        s = v.replace("\\", "\\\\").replace('"', '\\"').replace("\n", "\\n")
        return f'"{s}"'
    if isinstance(v, dict):
        return "{" + ",".join(f"{dump(k)}:{dump(x)}" for k, x in v.items()) + "}"
    if isinstance(v, list):
        return "[" + ",".join(dump(x) for x in v) + "]"
    raise TypeError(type(v))


def write(src: Path, dst: Path) -> None:
    data = yaml.load(src.read_text(encoding="utf-8"), Loader=yaml.BaseLoader)
    dst.write_text(dump(data) + "\n", encoding="utf-8")
    print(f"{dst.relative_to(ROOT)}")


def main() -> None:
    fixtures = ROOT / "tool" / "fixtures"
    for src in sorted(fixtures.glob("*.yaml")):
        write(src, src.with_suffix(".expected"))


if __name__ == "__main__":
    sys.exit(main())
