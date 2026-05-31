#!/usr/bin/env python3
"""
Install LLM dependencies for 'grepxcel draft', auto-detecting the best
available hardware (CUDA GPU → Metal → CPU-only).

Usage:
    python3 scripts/install_llm_deps.py

Use this instead of:
    pip install -r requirements-suggest.txt

The script installs huggingface_hub normally and compiles llama-cpp-python
with the right backend flags for the current machine. On machines with no
GPU the result is identical to a plain pip install — just CPU inference.
"""

import os
import subprocess
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _bootstrap import ensure_venv; ensure_venv()


def _run(cmd: list[str]) -> None:
    print(f'\n$ {" ".join(cmd)}', flush=True)
    result = subprocess.run(cmd)
    if result.returncode != 0:
        print(f'\nError: command failed with exit code {result.returncode}', file=sys.stderr)
        sys.exit(result.returncode)


def _detect_backend() -> tuple[str, dict]:
    """
    Detect best available compute backend.
    Returns (label, env_overrides) where env_overrides are merged into os.environ
    before calling pip to compile llama-cpp-python.
    """
    # ── CUDA (NVIDIA) ──────────────────────────────────────────────────────────
    try:
        result = subprocess.run(
            ['nvidia-smi', '--query-gpu=name,memory.total', '--format=csv,noheader'],
            capture_output=True, text=True, timeout=5,
        )
        if result.returncode == 0 and result.stdout.strip():
            gpus = result.stdout.strip().splitlines()
            for gpu in gpus:
                print(f'  Found NVIDIA GPU: {gpu.strip()}')
            # Check nvcc is available (CUDA Toolkit required for compilation)
            nvcc = subprocess.run(['nvcc', '--version'], capture_output=True, timeout=5)
            if nvcc.returncode != 0:
                print(
                    '\n  Warning: NVIDIA GPU detected but CUDA Toolkit (nvcc) not found.\n'
                    '  llama-cpp-python will be compiled for CPU only.\n'
                    '  To enable GPU acceleration, run: python3 scripts/install_cuda.py\n'
                    '  Then re-run this script.\n',
                    file=sys.stderr,
                )
            else:
                cmake_args = '-DGGML_CUDA=on'

                # Prefer the distro-packaged nvcc (/usr/bin/nvcc) over a manually
                # installed runfile toolkit (/usr/local/cuda-*/bin/nvcc) when both
                # exist. The distro nvcc is built against the system GCC and avoids
                # C++ stdlib compatibility issues on Ubuntu 24.04+.
                distro_nvcc = '/usr/bin/nvcc'
                if os.path.exists(distro_nvcc):
                    cmake_args += f' -DCMAKE_CUDA_COMPILER={distro_nvcc}'
                    print(f'  Using distro nvcc: {distro_nvcc} (avoids GCC compatibility issues)')
                else:
                    # CUDA 12.x only officially supports GCC ≤ 13.
                    # Ubuntu 24.04+ ships GCC 14+. Detect and add the override flag.
                    gcc_ver = subprocess.run(
                        ['gcc', '-dumpversion'], capture_output=True, text=True, timeout=5,
                    )
                    if gcc_ver.returncode == 0:
                        major = int(gcc_ver.stdout.strip().split('.')[0])
                        if major > 13:
                            print(f'  GCC {major} detected — CUDA 12.x supports up to GCC 13.')
                            print('  Adding -allow-unsupported-compiler flag automatically.')
                            cmake_args += ' -DCMAKE_CUDA_FLAGS=-allow-unsupported-compiler'

                return 'CUDA', {'CMAKE_ARGS': cmake_args}
    except (FileNotFoundError, subprocess.TimeoutExpired):
        pass

    # ── Metal (Apple Silicon / macOS) ─────────────────────────────────────────
    if sys.platform == 'darwin':
        try:
            result = subprocess.run(
                ['system_profiler', 'SPDisplaysDataType'],
                capture_output=True, text=True, timeout=5,
            )
            if 'Apple' in result.stdout or 'Metal' in result.stdout:
                print('  Found Apple Silicon / Metal GPU.')
                return 'Metal', {'CMAKE_ARGS': '-DGGML_METAL=on'}
        except (FileNotFoundError, subprocess.TimeoutExpired):
            pass

    # ── CPU fallback ───────────────────────────────────────────────────────────
    print('  No GPU detected — using CPU inference.')
    return 'CPU', {}


def main():
    print('Detecting hardware...')
    backend, cmake_env = _detect_backend()
    print(f'  Backend selected: {backend}\n')

    pip = sys.executable.replace('python3', 'pip').replace('python', 'pip')
    # Use the venv pip reliably
    pip = os.path.join(os.path.dirname(sys.executable), 'pip')

    # Install huggingface_hub and other non-compiled deps first
    _run([pip, 'install', 'huggingface_hub>=0.23'])

    # Install llama-cpp-python with the right compile flags
    env = {**os.environ, **cmake_env}
    cmd = [pip, 'install', 'llama-cpp-python>=0.2.90', '--force-reinstall', '--no-cache-dir']
    print(f'\n$ {" ".join(cmd)}')
    if cmake_env:
        for k, v in cmake_env.items():
            print(f'  (with {k}={v})')
    result = subprocess.run(cmd, env=env)
    if result.returncode != 0:
        print(f'\nError: llama-cpp-python installation failed.', file=sys.stderr)
        sys.exit(result.returncode)

    print(f'\nDone. llama-cpp-python installed with {backend} support.')
    print('You can now use: grepxcel draft data.xlsx')


if __name__ == '__main__':
    main()
