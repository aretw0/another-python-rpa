# -*- coding: utf-8 -*-

import pytest
from another_python_rpa.skeleton import fib

__author__ = "Arthur Aleksandro Alves Silva"
__copyright__ = "Arthur Aleksandro Alves Silva"
__license__ = "mit"


def test_fib():
    assert fib(1) == 1
    assert fib(2) == 1
    assert fib(7) == 13
    with pytest.raises(AssertionError):
        fib(-10)
