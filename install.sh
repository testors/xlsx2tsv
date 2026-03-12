#!/bin/sh

set -eu

usage() {
    cat <<'EOF'
Usage: ./install.sh [--prefix dir] [--bin-dir dir] [--cc compiler] [--cflags flags] [--ldflags flags]

Build xlsx2tsv with a portable release configuration and install it into a user-writable bin directory.

Options:
  --prefix dir    Install prefix. Default: $HOME/.local
  --bin-dir dir   Install directly into this bin directory
  --cc compiler   C compiler to use. Default: cc
  --cflags flags  Override C compiler flags
  --ldflags flags Override linker flags
  --help          Show this help
EOF
}

script_dir=$(
    CDPATH= cd -- "$(dirname -- "$0")" && pwd
)

prefix=${PREFIX:-"$HOME/.local"}
bin_dir=
cc=${CC:-cc}
cflags=${CFLAGS:-"-O3 -DNDEBUG -Wall -Wextra"}
ldflags=${LDFLAGS:-}
libs=${LIBS:-"-lz"}

while [ "$#" -gt 0 ]; do
    case "$1" in
        --prefix)
            [ "$#" -ge 2 ] || {
                echo "install.sh: --prefix requires a value" >&2
                exit 1
            }
            prefix=$2
            shift 2
            ;;
        --bin-dir)
            [ "$#" -ge 2 ] || {
                echo "install.sh: --bin-dir requires a value" >&2
                exit 1
            }
            bin_dir=$2
            shift 2
            ;;
        --cc)
            [ "$#" -ge 2 ] || {
                echo "install.sh: --cc requires a value" >&2
                exit 1
            }
            cc=$2
            shift 2
            ;;
        --cflags)
            [ "$#" -ge 2 ] || {
                echo "install.sh: --cflags requires a value" >&2
                exit 1
            }
            cflags=$2
            shift 2
            ;;
        --ldflags)
            [ "$#" -ge 2 ] || {
                echo "install.sh: --ldflags requires a value" >&2
                exit 1
            }
            ldflags=$2
            shift 2
            ;;
        --help|-h)
            usage
            exit 0
            ;;
        *)
            echo "install.sh: unknown option: $1" >&2
            usage >&2
            exit 1
            ;;
    esac
done

if [ -z "$bin_dir" ]; then
    bin_dir=$prefix/bin
fi

if ! command -v "$cc" >/dev/null 2>&1; then
    echo "install.sh: compiler not found: $cc" >&2
    exit 1
fi

mkdir -p "$bin_dir"

tmp_bin=$(mktemp "${TMPDIR:-/tmp}/xlsx2tsv.XXXXXX")
cleanup() {
    rm -f "$tmp_bin"
}
trap cleanup EXIT INT TERM

echo "Building xlsx2tsv with $cc..."
if ! "$cc" $cflags -o "$tmp_bin" \
    "$script_dir/xlsx_to_tsv.c" \
    "$script_dir/filter.c" \
    $ldflags $libs
then
    cat >&2 <<'EOF'
install.sh: build failed.
Make sure a C compiler is installed and zlib development files are available.
- macOS: install Xcode Command Line Tools
- Debian/Ubuntu: apt install build-essential zlib1g-dev
- Fedora/RHEL: dnf install gcc zlib-devel
EOF
    exit 1
fi

target=$bin_dir/xlsx_to_tsv
cp "$tmp_bin" "$target"
chmod 755 "$target"

echo "Installed: $target"

case ":${PATH:-}:" in
    *:"$bin_dir":*)
        ;;
    *)
        echo "Note: $bin_dir is not in PATH."
        echo "Add this line to your shell profile:"
        echo "  export PATH=\"$bin_dir:\$PATH\""
        ;;
esac
