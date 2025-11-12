from __future__ import absolute_import

# Copyright (c) 2010-2014 openpyxl
#
# Permission is hereby granted, free of charge, to any person obtaining a copy
# of this software and associated documentation files (the "Software"), to deal
# in the Software without restriction, including without limitation the rights
# to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
# copies of the Software, and to permit persons to whom the Software is
# furnished to do so, subject to the following conditions:
#
# The above copyright notice and this permission notice shall be included in
# all copies or substantial portions of the Software.
#
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
# THE SOFTWARE.
#
# @license: http://www.opensource.org/licenses/mit-license.php
import pytest

from openpyxl.styles import Style, Font, Color
from openpyxl.workbook import Workbook


def test_apply_vectorized_style_rows_shared():
    wb = Workbook()
    ws = wb.active

    font = Font()
    font.bold = True

    ws.apply_vectorized_style(rows=[1, 2], font=font)

    assert 1 in ws._styles and 2 in ws._styles
    assert ws._styles[1].font.bold is True
    assert ws._styles[1] is ws._styles[2]
    assert ws._styles[1].static is True
    assert 1 in ws.row_dimensions
    assert 2 in ws.row_dimensions


def test_apply_vectorized_style_cols_range_string():
    wb = Workbook()
    ws = wb.active

    font = Font()
    font.italic = True

    ws.apply_vectorized_style(cols='A:C', font=font)

    assert ws._styles['A'].font.italic is True
    assert ws._styles['A'] is ws._styles['C']
    assert ws._styles['B'] is ws._styles['A']
    assert 'A' in ws.column_dimensions
    assert 'C' in ws.column_dimensions


def test_apply_vectorized_style_non_shared():
    wb = Workbook()
    ws = wb.active

    base_style = Style()
    base_style.font = Font()
    base_style.font.underline = Font.UNDERLINE_SINGLE

    ws.apply_vectorized_style(rows='3:4', style=base_style, shared=False)

    assert ws._styles[3].font.underline == 'single'
    assert ws._styles[3] is not ws._styles[4]
    assert ws._styles[3].static is False
    assert base_style.font.underline == 'single'


def test_apply_vectorized_style_slice_semantics():
    wb = Workbook()
    ws = wb.active

    font = Font()
    font.color = Color(Color.RED)

    ws.apply_vectorized_style(rows=slice(5, 8), font=font)
    assert set(ws._styles) == {5, 6, 7}


def test_apply_vectorized_style_row_string_descending():
    wb = Workbook()
    ws = wb.active

    font = Font()
    font.bold = True

    ws.apply_vectorized_style(rows='5:3', font=font)

    assert set(ws._styles) == {5, 4, 3}


def test_apply_vectorized_style_column_iterable_normalization():
    wb = Workbook()
    ws = wb.active

    font = Font()
    font.strikethrough = True

    ws.apply_vectorized_style(cols=[1, '3', 'B'], font=font, shared=False)

    assert set(ws._styles) == {'A', 'B', 'C'}
    assert ws._styles['A'].font.strikethrough is True
    assert ws._styles['B'].font.strikethrough is True
    assert ws._styles['A'] is not ws._styles['B']


def test_apply_vectorized_style_invalid_usage():
    wb = Workbook()
    ws = wb.active

    with pytest.raises(ValueError):
        ws.apply_vectorized_style()

    with pytest.raises(ValueError):
        ws.apply_vectorized_style(rows=[1], cols=['A'])

    with pytest.raises(TypeError):
        ws.apply_vectorized_style(cols=object())

