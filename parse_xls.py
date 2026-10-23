"""Еженедельная статистика по DEV-сервисам.

Порядок работы:
1. Обновление внешних запросов в «Statistik Loader.xlsx» через Excel (COM).
2. Чтение листа «Quelldaten», удаление закрытых заявок и заявок не из очередей DEV.
3. Подсчёт заявок по сервисам [T000] DEV::*.
4. Сохранение результата (листы «Quelldaten» и «T000») в сетевую директорию
   под именем ДД.ММ.ГГГГ.xlsx.
5. Подготовка письма в Outlook (окно открывается для проверки перед отправкой).
"""

from datetime import date
from pathlib import Path

import pandas as pd

# Пути заданы через raw-строки: в обычных строках «\b» превращается в backspace,
# а «\'» экранирует кавычку (из-за этого исходный скрипт не компилировался).
QUELLE = Path(r'C:\Documents\Statistik Loader.xlsx')
ZIEL_DIR = Path(r'\\0x2srv\DFS-Abteilungen$\Developers\b-end')
ZIEL_DIR_LAUFWERK = r'S:\Developers\b-end'  # тот же путь через подключённый диск, для письма
EMPFAENGER = 'ruslan@0x22222222.com'

QUELL_BLATT = 'Quelldaten'
ERGEBNIS_BLATT = 'T000'
DEV_QUEUES = ['[1] Second Level::DEV', '[2] Third Level::DEV']
# Столбец I листа «Quelldaten» (9-й, индекс 8) — сервис заявки
DIENST_SPALTE_INDEX = 8

DIENST_PREFIX = '[T000] DEV::'
DIENSTE = ['CPP', 'JIT', 'JAVA-Entwicklung', 'Python', 'Logistik',
           'EDI-Mapping', 'EDI', 'B1-2', 'System', 'TP']

# calendar.month_name зависит от локали (обычно английская) — нужны немецкие названия
MONATE = ['Januar', 'Februar', 'März', 'April', 'Mai', 'Juni', 'Juli',
          'August', 'September', 'Oktober', 'November', 'Dezember']


def excel_aktualisieren(pfad: Path) -> None:
    """Обновляет все запросы книги и сохраняет её; Excel закрывается при любой ошибке."""
    import win32com.client

    xlapp = win32com.client.DispatchEx('Excel.Application')
    xlapp.DisplayAlerts = False
    try:
        wb = xlapp.Workbooks.Open(str(pfad))
        try:
            wb.RefreshAll()
            xlapp.CalculateUntilAsyncQueriesDone()
            wb.Save()
        finally:
            wb.Close(SaveChanges=False)
    finally:
        xlapp.Quit()


def daten_filtern(df: pd.DataFrame) -> pd.DataFrame:
    """Оставляет только открытые заявки из очередей DEV."""
    maske = (df['State'] != 'closed') & df['Queue'].isin(DEV_QUEUES)
    return df[maske]


def dienste_zaehlen(df: pd.DataFrame) -> pd.DataFrame:
    """Считает заявки по каждому сервису; сервисы без заявок получают 0."""
    namen = [DIENST_PREFIX + d for d in DIENSTE]
    spalte = df.iloc[:, DIENST_SPALTE_INDEX].astype(str)
    anzahl = spalte.value_counts().reindex(namen, fill_value=0)
    return pd.DataFrame({'Dienst': namen, 'Anzahl': anzahl.to_numpy()})


def speichern(df: pd.DataFrame, zaehlung: pd.DataFrame, ziel: Path) -> None:
    with pd.ExcelWriter(ziel, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name=QUELL_BLATT, index=False)
        zaehlung.to_excel(writer, sheet_name=ERGEBNIS_BLATT, index=False, header=False)


def mail_vorbereiten(dateiname: str, heute: date) -> None:
    import win32com.client

    woche = heute.isocalendar().week
    outlook = win32com.client.Dispatch('Outlook.Application')
    mail = outlook.CreateItem(0)
    mail.To = EMPFAENGER
    mail.Subject = f'Die Liste Woche #{woche:02d}, {MONATE[heute.month - 1]}'
    mail.HTMLBody = (f'Hallo,<br>die Datei {ZIEL_DIR_LAUFWERK}\\{dateiname} ist fertig.<br><br>'
                     'Viele Grüße<br>')
    mail.Display(True)


def main() -> None:
    print('Aktualisierung xlsx...')
    excel_aktualisieren(QUELLE)
    print('Aktualisierung ist fertig.')

    print('Lesen die xlsx Datei...')
    df = pd.read_excel(QUELLE, sheet_name=QUELL_BLATT, na_values=['NA'])

    print('Reinigung und Kalkulation...')
    df = daten_filtern(df)
    zaehlung = dienste_zaehlen(df)

    heute = date.today()
    dateiname = heute.strftime('%d.%m.%Y') + '.xlsx'
    ziel = ZIEL_DIR / dateiname
    speichern(df, zaehlung, ziel)
    print(f'Datei {ziel} ist gespeichert.')

    mail_vorbereiten(dateiname, heute)


if __name__ == '__main__':
    main()
