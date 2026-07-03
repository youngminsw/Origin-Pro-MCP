"""Shared fake Origin COM doubles.

Importable from both ``conftest.py`` (for the ``fake_origin`` fixture) and the
daemon/transport tests, which spin up several independent ``FakeOrigin``
instances on separate worker threads. Keeping the fakes here (instead of inside
``conftest``) lets every test module import the exact same COM surface.
"""


class FakePages:
    def __init__(self, pages):
        self._pages = pages

    @property
    def Count(self):
        return len(self._pages)

    def Item(self, i):
        return self._pages[i]


class FakeColumn:
    def __init__(self, name, col_type=0, long_name=""):
        self.Name = name
        self.Type = col_type  # COM designation: 0=Y, 2=Y Error, 3=X
        self.LongName = long_name


class FakeSheet:
    def __init__(self, name, columns=()):
        self.Name = name
        self.columns = list(columns)

    @property
    def Columns(self):
        return FakePages(self.columns)


class FakeBook:
    def __init__(self, name, sheets=("Sheet1",)):
        self.Name = name
        self.sheets = [
            s if isinstance(s, FakeSheet) else FakeSheet(s) for s in sheets
        ]

    @property
    def Layers(self):
        return FakePages(self.sheets)


class FakePlot:
    def __init__(self, name):
        self.Name = name
        self.activated = False

    def Activate(self):
        self.activated = True


class FakeLayer:
    def __init__(self, plot_names):
        self.DataPlots = FakePages([FakePlot(n) for n in plot_names])

    def Execute(self, script):
        return True


class FakeGraph:
    def __init__(self, name, plot_names=()):
        self.Name = name
        self.plot_names = list(plot_names)


class FakeMatrix:
    def __init__(self, name):
        self.Name = name


class FakeOrigin:
    """Mimics the Origin COM surface the tools rely on."""

    def __init__(self):
        self.books = [FakeBook("Book1")]
        self.graphs = [FakeGraph("Graph1")]
        self.execute_results = {}
        self.executed = []
        self.save_result = True
        self.load_result = True
        self.saved_paths = []
        self.put_result = True
        self.worksheet_data = ((1.0, 4.0), (2.0, 5.0))
        self.matrices = []
        self.matrix_data = {}
        self.lt_vars = {}

    @property
    def WorksheetPages(self):
        return FakePages(self.books)

    @property
    def GraphPages(self):
        return FakePages(self.graphs)

    @property
    def MatrixPages(self):
        return FakePages(self.matrices)

    def Execute(self, script):
        self.executed.append(script)
        for prefix, result in self.execute_results.items():
            if script.startswith(prefix):
                return result
        return True

    def FindWorksheet(self, target):
        for book in self.books:
            for j in range(book.Layers.Count):
                sheet = book.Layers.Item(j)
                if target == f"[{book.Name}]{sheet.Name}":
                    return sheet
        return None

    def FindGraphLayer(self, target):
        cache = self.__dict__.setdefault("_graph_layers", {})
        if target in cache:
            return cache[target]
        for graph in self.graphs:
            if target == f"[{graph.Name}]Layer1":
                cache[target] = FakeLayer(graph.plot_names)
                return cache[target]
        return None

    def FindMatrixSheet(self, target):
        for m in self.matrices:
            if target in (m.Name, f"[{m.Name}]MSheet1"):
                return m
        return None

    def PutMatrix(self, target, data):
        self.matrix_data[target] = [list(r) for r in data]
        return self.put_result

    def GetMatrix(self, target):
        if self.FindMatrixSheet(target) is None:
            return -2147352568
        grid = self.matrix_data.get(target, ((1.0, 2.0), (3.0, 4.0)))
        return tuple(tuple(r) for r in grid)

    def GetWorksheet(self, target):
        if self.FindWorksheet(target) is None:
            return -2147352568  # HRESULT int, as observed on Origin 2020
        return self.worksheet_data

    def PutWorksheet(self, target, data, row, col):
        return self.put_result

    def Save(self, path):
        self.saved_paths.append(path)
        return self.save_result

    def Load(self, path):
        return self.load_result

    def CreatePage(self, kind, name, template):
        return name

    def LTVar(self, name):
        return self.lt_vars.get(name, 0.0)

    def LTStr(self, name):
        return ""


class ThreadGuardedFake:
    """A COM-proxy stand-in that records ANY off-owner-thread attribute access.

    The watchdog and the heartbeat/idle monitor must NEVER dereference a COM
    proxy — they only ever deal in PIDs. We hand those threads the int pid (never
    this object) and assert ``touched`` stays False, proving they stayed
    COM-free. Defining no real methods/attributes here is deliberate: every
    access funnels through ``__getattr__`` so the owning-thread guard fires on
    each one (a real method would bypass ``__getattr__`` and defeat detection).
    """

    def __init__(self, owner_thread_id: int):
        import threading

        object.__setattr__(self, "_owner", owner_thread_id)
        object.__setattr__(self, "touched", False)

    def __getattr__(self, name):
        import threading

        if name in ("_owner", "touched"):
            return object.__getattribute__(self, name)
        if threading.get_ident() != object.__getattribute__(self, "_owner"):
            object.__setattr__(self, "touched", True)
            raise RuntimeError(f"COM proxy touched off-owner-thread: {name!r}")
        return object.__getattribute__(self, name)
