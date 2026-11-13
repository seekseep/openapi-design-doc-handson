"""Rough check for undeclared identifiers in VBA source (Option Explicit).

Not a parser: it strips strings and comments, collects declarations
(module-level, Dim, parameters, procedure names in all modules) and reports
identifiers that are none of those nor known VBA / Excel names.
"""
import re
import sys
from pathlib import Path

KNOWN = set("""
and as attribute boolean byref byte byval call case chr chr$ chrw clng close cstr date dim do each else elseif
end err exit explicit false for function get goto if iif in instr instrrev integer is lbound lcase left left$
len lof long loop ltrim ltrim$ mid mid$ mod new next not nothing object on open option optional or preserve
private property public redim replace resume right right$ rtrim rtrim$ select set space space$ split string
string$ sub then to trim trim$ true ubound ucase until variant vartype vbboolean vbbinarycompare vbcr vbcrlf
vblf vbtab vbobjecterror vbdirectory while with xor binary access read input output append freefile as
application dir error int abs isnumeric isempty isarray isobject typename format now year month day hour minute
second val cdbl cint clng cbool join filter array sgn fix hex oct asc ascw debug print raise number description
source mac grantaccesstomultiplefiles getopenfilename pathseparator const step me
msgbox vbinformation vbexclamation vbtextcompare vbstring strcomp mkdir rgb like thisworkbook paramarray
worksheet range collection hyperlink
""".split())


def strip(line: str) -> str:
    line = re.sub(r'"(?:[^"]|"")*"', '""', line)
    i = line.find("'")
    return line[:i] if i >= 0 else line


def join_continuations(lines):
    out = []
    buf = ""
    for l in lines:
        if l.rstrip().endswith(" _"):
            buf += l.rstrip()[:-2] + " "
            out.append("")
        else:
            out.append(buf + l)
            buf = ""
    return out


def main(paths):
    sources = {p: join_continuations([strip(l) for l in Path(p).read_text().splitlines()]) for p in paths}
    global_names = set()
    for p, lines in sources.items():
        global_names.add(Path(p).stem.lower())
        for l in lines:
            m = re.match(r"\s*(?:Public |Private )?(?:Function|Sub|Property Get|Property Let|Property Set)\s+(\w+)", l, re.I)
            if m:
                global_names.add(m.group(1).lower())
            m = re.match(r"\s*(?:Public|Private|Dim)\s+(?!Function|Sub|Property|Const)(.*)", l, re.I)
            if m and not l.startswith((" ", "\t")):
                for part in m.group(1).split(","):
                    n = re.match(r"\s*(\w+)", part)
                    if n:
                        global_names.add(n.group(1).lower())
            m = re.match(r"\s*(?:Public |Private )?Const\s+(\w+)", l, re.I)
            if m:
                global_names.add(m.group(1).lower())
    problems = 0
    for p, lines in sources.items():
        # Module-level declarations must come before the first procedure.
        seen_proc = False
        for no, l in enumerate(lines, 1):
            if re.match(r"(Public |Private )?(Function|Sub|Property)\b", l, re.I):
                seen_proc = True
            elif seen_proc and re.match(r"(Public|Private|Dim|Const)\s", l, re.I):
                print(f"{p}:{no}: module-level declaration after a procedure")
                problems += 1
        local = set()
        for no, l in enumerate(lines, 1):
            if re.match(r"\s*(Public |Private )?(Function|Sub|Property)\b", l, re.I):
                local = set()
                params = re.search(r"\((.*)\)", l)
                if params:
                    for part in params.group(1).split(","):
                        n = re.match(r"\s*(?:Optional\s+)?(?:ParamArray\s+)?(?:ByVal\s+|ByRef\s+)?(\w+)", part, re.I)
                        if n:
                            local.add(n.group(1).lower())
            for m in re.finditer(r"\b(?:Dim|ReDim(?: Preserve)?)\s+(.*)", l, re.I):
                for part in re.split(r",(?![^(]*\))", m.group(1)):
                    n = re.match(r"\s*(\w+)", part)
                    if n:
                        local.add(n.group(1).lower())
            if re.match(r"\s*(Attribute|#|Option|VERSION|BEGIN|END$|MultiUse)", l):
                continue
            code = re.sub(r"\.\s*\w+", ".", l)  # drop member names after "."
            code = re.sub(r"\bAs\s+(New\s+)?\w+", "", code, flags=re.I)
            code = re.sub(r"\b\w+:=", "", code)  # named arguments
            for tok in re.findall(r"(?<![\w&.])[A-Za-z_]\w*\$?", code):
                t = tok.lower()
                if re.match(r"(xl|vb)[A-Z]", tok):  # Excel / VBA constants
                    continue
                if t in KNOWN or t in local or t in global_names or t.rstrip("$") in KNOWN:
                    continue
                if re.match(r"\s*\w+:\s*$", l) or re.search(r"GoTo\s+" + tok, l):  # labels
                    continue
                print(f"{p}:{no}: {tok}")
                problems += 1
    return 1 if problems else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
