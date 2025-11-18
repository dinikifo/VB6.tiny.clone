import json
from copy import deepcopy


class VBJsonRuntime:
    SCHEMAS = {
        "Root": {},   # must be a dict, not a string
        "Customer": {
            "name": "",
            "age": 0
        }
    }

    @staticmethod
    def json_new(type_name: str):
        tpl = VBJsonRuntime.SCHEMAS.get(type_name)
        if tpl is None:
            # fallback: empty object
            return {}
        return deepcopy(tpl)

    @staticmethod
    def json_parse(text: str):
        return json.loads(text)

    @staticmethod
    def json_stringify(value):
        return json.dumps(value)

    @staticmethod
    def _split_path(path: str):
        # Parse "a.b[0].c" -> ["a","b",0,"c"]
        tokens = []
        buf = ""
        i = 0
        while i < len(path):
            ch = path[i]
            if ch == '.':
                if buf:
                    tokens.append(buf)
                    buf = ""
                i += 1
            elif ch == '[':
                if buf:
                    tokens.append(buf)
                    buf = ""
                i += 1
                num = ""
                while i < len(path) and path[i] != ']':
                    num += path[i]
                    i += 1
                i += 1  # skip ]
                tokens.append(int(num))
            else:
                buf += ch
                i += 1
        if buf:
            tokens.append(buf)
        return tokens

    @staticmethod
    def json_get(obj, path: str):
        parts = VBJsonRuntime._split_path(path)
        cur = obj
        for p in parts:
            if isinstance(p, int):
                cur = cur[p]
            else:
                cur = cur[p]
        return cur

    @staticmethod
    def json_set(obj, path: str, value):
        parts = VBJsonRuntime._split_path(path)
        if not isinstance(obj, (dict, list)):
            raise TypeError(
                f"JsonSet: first argument must be dict/list, got {type(obj).__name__}"
            )

        cur = obj
        for p in parts[:-1]:
            if isinstance(p, int):
                if not isinstance(cur, list):
                    raise TypeError(
                        f"JsonSet: trying to index non-list with [{p}], got {type(cur).__name__}"
                    )
                cur = cur[p]
            else:
                if not isinstance(cur, dict):
                    raise TypeError(
                        f"JsonSet: trying to use key '{p}' on non-dict {type(cur).__name__}"
                    )
                if p not in cur or cur[p] is None:
                    cur[p] = {}
                cur = cur[p]

        last = parts[-1]
        if isinstance(last, int):
            if not isinstance(cur, list):
                raise TypeError(
                    f"JsonSet: trying to index non-list with [{last}], got {type(cur).__name__}"
                )
            cur[last] = value
        else:
            if not isinstance(cur, dict):
                raise TypeError(
                    f"JsonSet: trying to use key '{last}' on non-dict {type(cur).__name__}"
                )
            cur[last] = value


class VBContext:
    def __init__(self):
        self.globals = {}

    def set_var(self, name, value):
        self.globals[name] = value

    def get_var(self, name):
        return self.globals.get(name, None)

    def has_var(self, name):
        return name in self.globals
