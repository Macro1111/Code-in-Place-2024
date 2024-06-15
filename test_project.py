import pytest

from project import nombre_nuevo_archivo, consecutivo, numero_letras

def test_numeros_letras():
    assert numero_letras(101) == "CIENTO UNO"
    assert numero_letras(1100) == "MIL CIEN"
    assert numero_letras(100000) == "CIEN MIL"
    assert numero_letras(1000000) == "UN MILLÓN"

def test_consecutivo():
    assert consecutivo(3) == 303
    assert consecutivo(10) == 310
    assert consecutivo(100) == 400

def test_probar_nombre():
    assert nombre_nuevo_archivo("archivo") == "archivo.docx"
    assert nombre_nuevo_archivo("Prueba") == "Prueba.docx"
    assert nombre_nuevo_archivo("1. Cuenta de Cobro") == "1. Cuenta de Cobro.docx"