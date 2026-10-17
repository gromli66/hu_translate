# -*- coding: utf-8 -*-
"""PDF из готового DOCX через LibreOffice без окна: обновить оглавления, список рисунков и поля (SEQ, PAGE),
напечатать PDF. Решение Максима 16.10: в PDF «на месте» не помещающийся русский текст ужимался — плавающий кегль;
PDF из DOCX набран как документ, написанный по-русски.

Запускается интерпретатором с модулем uno (системный python3 в образе, python LibreOffice на Windows), не python
приложения:  python3 lo_pdf.py вход.docx выход.pdf   — код выхода 0 и готовый PDF, иначе причина в stderr.
Выход .docx — сохранение в DOCX без обновления полей (старый .doc заказчика → DOCX для перевода).
Обычный «soffice --convert-to pdf» оглавление не обновляет — в PDF осталась бы заглушка поля."""
import os
import shutil
import subprocess
import sys
import tempfile
import time

import uno
from com.sun.star.beans import PropertyValue
from com.sun.star.connection import NoConnectException
from com.sun.star.lang import DisposedException


def _prop(name, value):
    p = PropertyValue()
    p.Name, p.Value = name, value
    return p


def convert(src, dst, soffice="soffice"):
    port = 20000 + os.getpid() % 20000
    profile = tempfile.mkdtemp(prefix="lo_profile_")
    url = f"socket,host=127.0.0.1,port={port};urp;StarOffice.ComponentContext"
    proc = subprocess.Popen([soffice, "--headless", "--invisible", "--nologo", "--norestore", "--nodefault",
                             "-env:UserInstallation=" + uno.systemPathToFileUrl(profile), "--accept=" + url],
                            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    try:
        local = uno.getComponentContext()
        resolver = local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local)
        ctx = None
        for _ in range(240):                                # первый запуск с чистым профилем — до ~1 мин
            try:
                ctx = resolver.resolve("uno:" + url)
                break
            except NoConnectException:
                time.sleep(0.5)
        if ctx is None:
            raise RuntimeError("LibreOffice не ответил за 2 мин")
        desktop = ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
        doc = desktop.loadComponentFromURL(uno.systemPathToFileUrl(os.path.abspath(src)), "_blank", 0,
                                           (_prop("Hidden", True),))
        if doc is None:                                     # LibreOffice не открыл файл (бывает разово) — понятной ошибкой
            raise RuntimeError(f"LibreOffice не открыл {os.path.basename(src)}")
        try:
            if dst.lower().endswith(".docx"):      # .doc → DOCX; без return — LibreOffice ниже закрывается и тут
                doc.storeToURL(uno.systemPathToFileUrl(os.path.abspath(dst)), (_prop("FilterName", "MS Word 2007 XML"),))
            else:
                idx = doc.getDocumentIndexes()
                for _ in range(2):          # второй проход — номера страниц после того, как оглавление заняло место
                    for i in range(idx.getCount()):
                        idx.getByIndex(i).update()
                    doc.getTextFields().refresh()
                doc.storeToURL(uno.systemPathToFileUrl(os.path.abspath(dst)), (_prop("FilterName", "writer_pdf_Export"),))
        finally:
            doc.close(True)
        try:
            desktop.terminate()
        except DisposedException:                           # LibreOffice закрылся раньше, чем ответил, — это и нужно
            pass
    finally:
        try:
            proc.wait(30)
        except subprocess.TimeoutExpired:
            proc.kill()
        shutil.rmtree(profile, ignore_errors=True)


if __name__ == "__main__":
    convert(sys.argv[1], sys.argv[2], *(sys.argv[3:4] or []))
