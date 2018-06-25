import re

pat = re.compile(r"^((\d+)|(\d+\.\d+)|(\d+\.\d+\.\d+)|(\d+\.\d+\.\d+\.\d+))\s(.*)")
st = "1.2.3 Hello"
mat = pat.match(st)
print(mat.group(1))