python -m pip download -d libs \
--only-binary=:all: \
--platform manylinux2014_x86_64  \
--python-version 3.7.10 \
--implementation cp \
-r requirements.txt
