import pytest

def test_addition():
    assert 1 + 1 == 2

def test_string():
    assert "hello".upper() == "HELLO"

class TestClass:
    def test_one(self):
        x = "this"
        assert "h" in x
