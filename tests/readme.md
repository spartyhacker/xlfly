
using testpy
https://packaging.python.org/en/latest/guides/using-testpypi/

change to it
pip config set global.index-url https://test.pypi.org/simple

change back
pip config set global.index-url https://pypi.org/simple

set up xlfly icon
py -3.12 -m xlfly.scripts --init -t D:\git\xlfly2\xlfly\tests\templates