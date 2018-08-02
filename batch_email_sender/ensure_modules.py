# -*- coding: utf-8 -*-
import importlib.util

modules_to_check = [
    'PyQt5',
    'openpyxl',
    'keyring'
]

if not all(importlib.util.find_spec(name) for name in modules_to_check):
    import ensurepip

    ensurepip.bootstrap(upgrade=False, user=True)
    import pip

    for name in modules_to_check:
        pip.main(['install', "--user", name])
