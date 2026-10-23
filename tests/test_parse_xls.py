"""Тесты логики обработки без Excel и Outlook (COM в CI недоступен)."""

from datetime import date

import pandas as pd

import parse_xls as p


def _quelldaten(zeilen):
    """Строит таблицу, в которой сервис находится в столбце I, как в «Statistik Loader.xlsx»."""
    spalten = list('ABCDEFGH')
    daten = []
    for state, queue, dienst in zeilen:
        zeile = dict.fromkeys(spalten, 0)
        zeile['Dienst'] = dienst
        zeile['State'] = state
        zeile['Queue'] = queue
        daten.append(zeile)
    return pd.DataFrame(daten, columns=spalten + ['Dienst', 'State', 'Queue'])


DATEN = _quelldaten([
    ('open', '[1] Second Level::DEV', '[T000] DEV::CPP'),
    ('open', '[2] Third Level::DEV', '[T000] DEV::CPP'),
    ('closed', '[1] Second Level::DEV', '[T000] DEV::CPP'),
    ('open', '[1] Second Level::OPS', '[T000] DEV::Python'),
    ('open', '[2] Third Level::DEV', '[T000] DEV::TP'),
    ('open', '[2] Third Level::DEV', 'Sonstiges'),
])


def test_filter_entfernt_geschlossene_und_fremde_queues():
    gefiltert = p.daten_filtern(DATEN)
    assert len(gefiltert) == 4
    assert (gefiltert['State'] != 'closed').all()
    assert gefiltert['Queue'].isin(p.DEV_QUEUES).all()


def test_zaehlung_enthaelt_alle_dienste():
    zaehlung = p.dienste_zaehlen(p.daten_filtern(DATEN))
    werte = dict(zip(zaehlung['Dienst'], zaehlung['Anzahl']))
    assert len(werte) == len(p.DIENSTE)
    assert werte['[T000] DEV::CPP'] == 2
    assert werte['[T000] DEV::TP'] == 1
    assert werte['[T000] DEV::Python'] == 0
    assert 'Sonstiges' not in werte


def test_speichern_legt_beide_blaetter_an(tmp_path):
    gefiltert = p.daten_filtern(DATEN)
    ziel = tmp_path / 'out.xlsx'
    p.speichern(gefiltert, p.dienste_zaehlen(gefiltert), ziel)
    with pd.ExcelFile(ziel) as xls:
        assert xls.sheet_names == [p.QUELL_BLATT, p.ERGEBNIS_BLATT]
        t000 = pd.read_excel(xls, p.ERGEBNIS_BLATT, header=None)
    assert t000.shape == (len(p.DIENSTE), 2)


def test_deutsche_monatsnamen():
    assert p.MONATE[date(2026, 10, 23).month - 1] == 'Oktober'
