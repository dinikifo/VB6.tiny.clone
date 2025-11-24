import re
from runtime import VBJsonRuntime, VBContext
import runtime as vb_runtime


class VBInterpreter:
    """
    VB-like interpreter:
    - Subs without parameters.
    - Functions without parameters, returning via assigning to function name.
    - Dim, assignment.
    - Expressions:
        * string literals (VB-style "" -> ")
        * numeric literals
        * variables, object.property
        * function calls (built-ins + user-defined no-arg functions)
        * string concatenation with &
        * numeric addition with +
    - If ... Then ... Else ... End If (with nesting).
    - Loops:
        * While <cond> ... Wend
        * Do While <cond> ... Loop
        * Do ... Loop While <cond>
    - Built-ins:
        MsgBox
        JsonNew, JsonParse, JsonStringify, JsonGet, JsonSet
        BrowserEvalJs
    """

    def __init__(self, ctx: VBContext):
        self.ctx = ctx
        self.procedures = {}   # name -> list of lines (Subs)
        self.functions = {}    # name -> list of lines (Functions)
        self._msgbox = lambda text: print(f"[MsgBox] {text}")  # patched by GUI

    # ------------ Helpers ------------ #

    def _normalize_lines(self, body: str):
        """
        Handle VB-style line continuation: lines ending with '_' are joined
        with the following line. Also strips comments-only and empty lines.
        """
        raw_lines = body.splitlines()
        logical_lines = []
        acc = ""

        for raw_line in raw_lines:
            line = raw_line.rstrip()

            # skip empty lines if not in continuation
            if not acc and line.strip() == "":
                continue

            stripped = line.lstrip()
            if not acc and stripped.startswith("'"):
                continue

            if line.endswith("_"):
                part = line[:-1].rstrip()
                if acc:
                    acc += " " + part
                else:
                    acc = part
                continue
            else:
                part = line
                if acc:
                    acc += " " + part
                    full = acc.strip()
                    acc = ""
                else:
                    full = part.strip()

                if full and not full.startswith("'"):
                    logical_lines.append(full)

        if acc:
            full = acc.strip()
            if full and not full.startswith("'"):
                logical_lines.append(full)

        return logical_lines

    # ------------ Parsing source ------------ #

    def load_source(self, source: str):
        """
        Parse Subs and Functions from VB-like source.
        """

        # Subs
        sub_pattern = re.compile(
            r"Sub\s+(\w+)\s*\(\)\s*(.*?)End Sub",
            re.IGNORECASE | re.DOTALL
        )
        for m in sub_pattern.finditer(source):
            name = m.group(1)
            body = m.group(2)
            lines = self._normalize_lines(body)
            self.procedures[name] = lines

        # Functions (no parameters for now)
        func_pattern = re.compile(
            r"Function\s+(\w+)\s*\(\)\s*(.*?)End Function",
            re.IGNORECASE | re.DOTALL
        )
        for m in func_pattern.finditer(source):
            name = m.group(1)
            body = m.group(2)
            lines = self._normalize_lines(body)
            self.functions[name] = lines

    # ------------ Running procedures & functions ------------ #

    def call_sub(self, name: str):
        lines = self.procedures.get(name)
        if not lines:
            return
        self.run_lines(lines, 0, len(lines))

    def call_function(self, name: str):
        lines = self.functions.get(name)
        if not lines:
            print(f"[VB] Function {name} not found")
            return None
        self.ctx.set_var(name, None)
        self.run_lines(lines, 0, len(lines))
        return self.ctx.get_var(name)

    def run_lines(self, lines, start, end):
        i = start
        while i < end:
            line = lines[i].strip()
            upper = line.upper()

            if upper.startswith("IF ") and " THEN" in upper:
                i = self._exec_if_block(lines, i, end)
                continue

            if upper.startswith("WHILE "):
                i = self._exec_while_block(lines, i, end)
                continue

            if upper.startswith("DO"):
                i = self._exec_do_block(lines, i, end)
                continue

            self.exec_line(line)
            i += 1

    # ------------ If / Else ------------ #

    def _coerce_for_compare(self, left, right):
        def to_num(x):
            if isinstance(x, (int, float)):
                return x, True
            if isinstance(x, str):
                try:
                    if "." in x:
                        return float(x), True
                    else:
                        return int(x), True
                except ValueError:
                    return x, False
            return x, False

        l_num, l_is = to_num(left)
        r_num, r_is = to_num(right)

        if l_is and r_is:
            return l_num, r_num

        return left, right

    def _eval_condition(self, cond: str) -> bool:
        cond = cond.strip()
        if cond == "":
            return False

        m = re.match(r"^(.*?)(=|<>|<=|>=|<|>)(.*)$", cond)
        if m:
            left_text = m.group(1).strip()
            op = m.group(2)
            right_text = m.group(3).strip()

            left_val = self.eval_expr(left_text)
            right_val = self.eval_expr(right_text)

            left_val, right_val = self._coerce_for_compare(left_val, right_val)

            try:
                if op == "=":
                    return left_val == right_val
                elif op == "<>":
                    return left_val != right_val
                elif op == "<":
                    return left_val < right_val
                elif op == ">":
                    return left_val > right_val
                elif op == "<=":
                    return left_val <= right_val
                elif op == ">=":
                    return left_val >= right_val
            except Exception as e:
                print(f"[VB] Error comparing values in If: {e}")
                return False

        val = self.eval_expr(cond)
        return bool(val)

    def _exec_if_block(self, lines, i, end):
        line = lines[i]
        uline = line.upper()
        then_pos = uline.rfind("THEN")
        cond_text = line[2:then_pos].strip()  # after "If" up to "Then"

        then_start = i + 1
        depth = 1
        else_index = None
        j = i + 1

        while j < end and depth > 0:
            cur = lines[j].strip()
            ucur = cur.upper()
            if ucur.startswith("IF ") and " THEN" in ucur:
                depth += 1
            elif ucur == "END IF":
                depth -= 1
            elif ucur == "ELSE" and depth == 1 and else_index is None:
                else_index = j
            j += 1

        end_if = j - 1
        if depth != 0:
            print("[VB] Warning: unmatched If/End If")
            return j

        then_end = else_index if else_index is not None else end_if

        else_start = None
        if else_index is not None:
            else_start = else_index + 1
            else_end = end_if
        else:
            else_end = None

        cond_value = self._eval_condition(cond_text)

        if cond_value:
            self.run_lines(lines, then_start, then_end)
        else:
            if else_index is not None:
                self.run_lines(lines, else_start, else_end)

        return end_if + 1

    # ------------ While / Wend ------------ #

    def _exec_while_block(self, lines, i, end):
        line = lines[i]
        cond_text = line[5:].strip()  # after "While"

        depth = 1
        j = i + 1
        while j < end and depth > 0:
            cur = lines[j].strip().upper()
            if cur.startswith("WHILE "):
                depth += 1
            elif cur == "WEND":
                depth -= 1
            j += 1

        if depth != 0:
            print("[VB] Warning: unmatched While/Wend")
            return j

        block_start = i + 1
        block_end = j - 1

        while self._eval_condition(cond_text):
            self.run_lines(lines, block_start, block_end)

        return j

    # ------------ Do / Loop ------------ #

    def _exec_do_block(self, lines, i, end):
        line = lines[i].strip()
        uline = line.upper()

        # Do While <cond> (pre-test)
        if uline.startswith("DO WHILE"):
            cond_text = line[8:].strip()

            depth = 1
            j = i + 1
            while j < end and depth > 0:
                cur = lines[j].strip().upper()
                if cur.startswith("DO"):
                    depth += 1
                elif cur.startswith("LOOP"):
                    depth -= 1
                j += 1

            if depth != 0:
                print("[VB] Warning: unmatched Do/Loop")
                return j

            block_start = i + 1
            block_end = j - 1

            while self._eval_condition(cond_text):
                self.run_lines(lines, block_start, block_end)

            return j

        # Do ... Loop While <cond> (post-test)
        else:
            depth = 1
            j = i + 1
            loop_index = None

            while j < end and depth > 0:
                cur = lines[j].strip().upper()
                if cur.startswith("DO"):
                    depth += 1
                elif cur.startswith("LOOP"):
                    depth -= 1
                    if depth == 0:
                        loop_index = j
                        break
                j += 1

            if loop_index is None:
                print("[VB] Warning: unmatched Do/Loop")
                return j

            loop_line = lines[loop_index].strip()
            uloop = loop_line.upper()

            cond_text = "False"
            if uloop.startswith("LOOP WHILE"):
                cond_text = loop_line[10:].strip()
            else:
                print("[VB] Only 'Loop While <cond>' supported (post-test)")

            block_start = i + 1
            block_end = loop_index

            while True:
                self.run_lines(lines, block_start, block_end)
                if not self._eval_condition(cond_text):
                    break

            return loop_index + 1

    # ------------ Simple statements ------------ #

    def _find_assignment_equals(self, line: str) -> int:
        """Find the position of an '=' that is not inside a string literal.

        This prevents lines like:
            lstPostings.Add "New journal created: ID=" & x
        from being misinterpreted as assignments.
        """
        in_string = False
        prev = ""
        for i, ch in enumerate(line):
            if ch == '"' and prev != '\\':
                in_string = not in_string
            if ch == "=" and not in_string:
                return i
            prev = ch
        return -1

    def exec_line(self, line: str):
        upper = line.upper()
        if upper.startswith("DIM "):
            self._exec_dim(line[3:].strip())
            return

        eq_pos = self._find_assignment_equals(line) if "=" in line else -1
        if eq_pos != -1 and not upper.startswith("IF "):
            left = line[:eq_pos].strip()
            right = line[eq_pos + 1 :].strip()
            value = self.eval_expr(right)
            self.assign(left, value)
            return

        self.exec_call(line)

    def _exec_dim(self, decl: str):
        names = [n.strip() for n in decl.split(",") if n.strip()]
        for name in names:
            base = re.split(r'\s+', name)[0]
            if not self.ctx.has_var(base):
                self.ctx.set_var(base, None)

    def assign(self, left: str, value):
        if "." in left:
            var_name, attr = map(str.strip, left.split(".", 1))
            obj = self.ctx.get_var(var_name)
            if obj is None:
                obj = {}
                self.ctx.set_var(var_name, obj)
            if hasattr(obj, attr):
                setattr(obj, attr, value)
            else:
                try:
                    obj[attr] = value
                except TypeError:
                    print(f"[VB] Cannot assign {attr} on {obj}")
        else:
            self.ctx.set_var(left, value)

    # ------------ Expressions ------------ #

    def eval_expr(self, expr: str):
        expr = expr.strip()
        if expr == "":
            return ""

        # 0) If it's a plain string literal, let eval_atom handle it
        #    This avoids treating "1+2" or "A&B" as operators.
        if expr.startswith('"') and expr.endswith('"'):
            return self.eval_atom(expr)

        # 1) String concatenation with &
        if "&" in expr:
            parts = [p.strip() for p in expr.split("&")]
            values = [self.eval_atom(p) for p in parts]
            return "".join("" if v is None else str(v) for v in values)

        # 2) Single binary arithmetic op: +, -, *, /
        #    Pattern:  left op right   (no chaining)
        m = re.match(r"^(.*?)([+\-*/])(.*)$", expr)
        if m:
            left_text = m.group(1).strip()
            op = m.group(2)
            right_text = m.group(3).strip()

            # if either side is empty, treat as a plain atom (e.g. "-5" not supported yet)
            if left_text != "" and right_text != "":
                left_val = self.eval_atom(left_text)
                right_val = self.eval_atom(right_text)

                try:
                    a = float(left_val)
                except (TypeError, ValueError):
                    a = 0.0
                try:
                    b = float(right_val)
                except (TypeError, ValueError):
                    b = 0.0

                if op == "+":
                    res = a + b
                elif op == "-":
                    res = a - b
                elif op == "*":
                    res = a * b
                elif op == "/":
                    res = a / b if b != 0 else 0.0
                else:
                    res = 0.0

                # return int when exact
                if isinstance(res, float) and res.is_integer():
                    return int(res)
                return res

        # 3) Fallback: single atom (variable, literal, function call...)
        return self.eval_atom(expr)


    def eval_atom(self, atom: str):
        atom = atom.strip()
        if atom == "":
            return ""

        # string literal – VB-style: "" -> "
        if atom.startswith('"') and atom.endswith('"'):
            inner = atom[1:-1]
            inner = inner.replace('""', '"')
            return inner

        # numeric literal
        if re.fullmatch(r"\d+(\.\d+)?", atom):
            if "." in atom:
                return float(atom)
            else:
                return int(atom)

        # function call
        func_match = re.match(r"(\w+)\s*\((.*)\)$", atom)
        if func_match:
            name = func_match.group(1)
            arg_str = func_match.group(2).strip()
            args = self._parse_args(arg_str) if arg_str else []
            eval_args = [self.eval_expr(a) for a in args]
            return self._call_function(name, eval_args)

        # variable.property
        if "." in atom:
            var_name, attr = map(str.strip, atom.split(".", 1))
            obj = self.ctx.get_var(var_name)
            if obj is None:
                return None
            if hasattr(obj, attr):
                return getattr(obj, attr)
            try:
                return obj[attr]
            except Exception:
                return None

        # bare variable
        return self.ctx.get_var(atom)

    def _parse_args(self, arg_str: str):
        args = []
        buf = []
        in_string = False
        depth = 0
        prev = ""
        for ch in arg_str:
            if ch == '"' and prev != '\\':
                in_string = not in_string
            if not in_string:
                if ch == '(':
                    depth += 1
                elif ch == ')':
                    depth -= 1
                elif ch == ',' and depth == 0:
                    arg = "".join(buf).strip()
                    if arg:
                        args.append(arg)
                    buf = []
                    prev = ch
                    continue
            buf.append(ch)
            prev = ch
        if buf:
            args.append("".join(buf).strip())
        return args

    def _call_function(self, name: str, args):
        upper = name.upper()

        # JSON built-ins
        if upper == "JSONNEW":
            return VBJsonRuntime.json_new(str(args[0]))
        if upper == "JSONPARSE":
            return VBJsonRuntime.json_parse(str(args[0]))
        if upper == "JSONSTRINGIFY":
            return VBJsonRuntime.json_stringify(args[0])
        if upper == "JSONGET":
            obj, path = args[0], str(args[1])
            return VBJsonRuntime.json_get(obj, path)

        # Accounting built-in: NewJournal(date, description, [period])
        if upper == "NEWJOURNAL":
            app_data = self.ctx.get_var("AppData")
            date = str(args[0]) if len(args) > 0 else ""
            desc = str(args[1]) if len(args) > 1 else ""
            period = None
            if len(args) > 2 and args[2] is not None:
                period = str(args[2])
            return vb_runtime.create_journal(app_data, date, desc, period)

        print(f"[VB] Unknown function in expression: {name}")
        return None

    # ------------ Calls as statements ------------ #

    def exec_call(self, line: str):
        # Method call: Obj.Method arg1, arg2, ...
        m_method = re.match(r"(\w+)\.(\w+)\s*(.*)", line)
        if m_method:
            obj_name = m_method.group(1)
            meth_name = m_method.group(2)
            arg_str = m_method.group(3).strip()
            if arg_str:
                if arg_str.startswith("(") and arg_str.endswith(")"):
                    arg_str = arg_str[1:-1]
                args = self._parse_args(arg_str)
                eval_args = [self.eval_expr(a) for a in args]
            else:
                eval_args = []

            obj = self.ctx.get_var(obj_name)
            if obj is None:
                print(f"[VB] Method call on unknown object: {obj_name}")
                return
            if not hasattr(obj, meth_name):
                print(f"[VB] Object {obj_name} has no method {meth_name}")
                return
            try:
                getattr(obj, meth_name)(*eval_args)
            except Exception as e:
                print(f"[VB] Error calling {obj_name}.{meth_name}: {e}")
            return

        # Bare procedure / built-in: Name arg1, arg2, ...
        m = re.match(r"(\w+)\s*(.*)", line)
        if not m:
            print(f"[VB] Cannot parse line: {line}")
            return
        name = m.group(1)
        arg_str = m.group(2).strip()
        if arg_str:
            if arg_str.startswith("(") and arg_str.endswith(")"):
                arg_str = arg_str[1:-1]
            args = self._parse_args(arg_str)
            eval_args = [self.eval_expr(a) for a in args]
        else:
            eval_args = []

        upper = name.upper()

        if upper == "MSGBOX":
            text = eval_args[0] if eval_args else ""
            self._msgbox(text)
            return

        if upper == "JSONSET":
            if len(eval_args) < 3:
                print("[VB] JsonSet requires 3 args")
                return
            obj, path, val = eval_args[0], str(eval_args[1]), eval_args[2]
            VBJsonRuntime.json_set(obj, path, val)
            return

        if upper == "BROWSEREVALJS":
            if len(eval_args) < 2:
                print("[VB] BrowserEvalJs requires browser and script")
                return
            browser_obj, script = eval_args[0], eval_args[1]
            if hasattr(browser_obj, "EvalJs"):
                try:
                    browser_obj.EvalJs(script)
                except Exception as e:
                    print(f"[VB] Error in BrowserEvalJs: {e}")
            else:
                print("[VB] First argument to BrowserEvalJs is not a browser object")
            return

        # Accounting built-in: PostEntry accountCode, assetTypeCode, period, journalId, amount
        if upper == "POSTENTRY":
            if len(eval_args) < 5:
                print("[VB] PostEntry requires 5 args: accountCode, assetTypeCode, period, journalId, amount")
                return
            app_data = self.ctx.get_var("AppData")
            account_code = str(eval_args[0])
            asset_type_code = str(eval_args[1])
            period = str(eval_args[2])
            try:
                journal_id = int(eval_args[3])
            except (TypeError, ValueError):
                journal_id = 0
            try:
                amount = float(eval_args[4])
            except (TypeError, ValueError):
                amount = 0.0
            vb_runtime.post_entry(app_data, account_code, asset_type_code, period, journal_id, amount)
            return

        print(f"[VB] Unknown statement: {line}")
