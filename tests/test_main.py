"""Tests d'exemple."""

import pytest


def test_exemple():
    """Test d'exemple."""
    assert 1 + 1 == 2


def test_import():
    """Test que el paquet es pot importar."""
    import long_descriptions
    
    assert long_descriptions.__version__ == "0.1.0"
