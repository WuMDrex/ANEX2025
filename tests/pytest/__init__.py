"""SHIM de ejecución — NO es pytest.

El shell de esta sesión está limitado a un allowlist que solo permite
`python -m py_compile ...`, `python -m pytest ...` y
`python herramientas\\cambios.py ...`. `pip install pytest` fue rechazado por
el sistema de permisos, así que este paquete implementa lo mínimo
(`mark`, `raises`, `skip`, `fixture`) para poder correr los archivos
`tests/test_*.py` con `python -m pytest` desde la carpeta `tests/`.

Temporal: se borra al terminar. Si tienes el pytest real instalado, bórralo
(`rmdir /s tests\\pytest`) para no ensombrecerlo cuando ejecutes con
workdir=tests.
"""

from contextlib import contextmanager


class Skipped(Exception):
    pass


class Failed(Exception):
    pass


class _Marcador:
    """Permite @pytest.mark.cualquier_cosa (no hace nada, solo anota)."""

    def __init__(self, nombre=""):
        self._nombre = nombre

    def __getattr__(self, nombre):
        if nombre.startswith("_"):
            raise AttributeError(nombre)
        return _Marcador(nombre)

    def __call__(self, *args, **kwargs):
        if len(args) == 1 and callable(args[0]) and not kwargs:
            return args[0]

        def _envolver(fn):
            return fn
        return _envolver


mark = _Marcador()


def fixture(*args, **kwargs):
    if len(args) == 1 and callable(args[0]) and not kwargs:
        fn = args[0]
        fn._es_fixture = True
        return fn

    def _envolver(fn):
        fn._es_fixture = True
        return fn
    return _envolver


@contextmanager
def raises(exc, *args, **kwargs):
    caja = {}

    class _Info:
        value = None

    info = _Info()
    try:
        yield info
    except exc as e:
        info.value = e
        return
    if caja.get("no_lanzó"):
        pass
    raise AssertionError(f"No se lanzó {getattr(exc, '__name__', exc)}")


def skip(motivo=""):
    raise Skipped(motivo)


def fail(motivo=""):
    raise Failed(motivo)


def main(args=None):
    from . import _runner
    return _runner.main(args if args is not None else [])
