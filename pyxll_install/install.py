import winreg as wr
import re


def set_options(path):
    opts_path = r"Software\Microsoft\Office\16.0\Excel\Options"
    key = wr.OpenKey(wr.HKEY_CURRENT_USER, opts_path, 0, wr.KEY_SET_VALUE)

    open_key = []
    i = 1
    while True:
        try:
            name, _, _ = wr.EnumValue(key, i)
        except WindowsError:
            break
        if 'OPEN' in name:
            open_key.append(name)
        i += 1

    try:
        last_open_key = sorted(open_key, reverse=True)[0]
        pos = start_pos(last_open_key)
        new_key = last_open_key[:-pos] + str(int(last_open_key[-pos:]) + 1)
    except IndexError:
        new_key = 'OPEN'

    wr.SetValueEx(key, new_key, 0, wr.REG_SZ, f"/R {path}")
    wr.CloseKey(key)


def set_add_in_key(path):
    add_in_path = r"Software\Microsoft\Office\16.0\Excel\Add-in Manager"
    key = wr.OpenKey(wr.HKEY_CURRENT_USER, add_in_path, 0, wr.KEY_SET_VALUE)
    wr.SetValueEx(key, f"{path}", 0, wr.REG_SZ, f"{path}")
    wr.CloseKey(key)


def start_pos(string):
    for match in re.finditer(r'\d+', string):
        start_pos = match.span()
    return start_pos[0]


def main(path):
    set_options(path)
    set_add_in_key(path)


def add_pyxll_registry_keys(path):
    set_options(path)
    set_add_in_key(path)


if __name__ == "__main__":
    import os
    import pathlib

    base_pyxll_path = pathlib.Path(os.getenv("USERPROFILE"))
    pyxll_path = base_pyxll_path / '.pyxll' / 'pyxll.xll'
    assert os.path.exists(pyxll_path)
    main(pyxll_path)
