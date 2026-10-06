# -*- coding: utf-8 -*-
"""Команды администратора веб-сервиса (в контейнере: docker compose exec hut python manage.py …).

  python manage.py user add <логин> [--name "Иванов М."] [--role user|expert|admin] [--password …]
  python manage.py user role <логин> <роль>
  python manage.py user passwd <логин> [--password …]
  python manage.py user disable <логин> | enable <логин>
  python manage.py user list
  python manage.py backup                       — копия базы сейчас (сервис делает её и сам раз в сутки)
  python manage.py project add <имя> <файл>     — новый проект (комплект) с глоссарием заказчика (md, docx, xlsx)
  python manage.py project list
Без --password пароль спрашивается с клавиатуры (так он не остаётся в истории команд)."""
import re
import sys
import json
import time
import getpass
import argparse

from server import db, security

ROLES = ("user", "expert", "admin")


def ask_password(given):
    if given:
        return given
    p1 = getpass.getpass("Пароль (не короче 8 знаков): ")
    if p1 != getpass.getpass("Ещё раз: "):
        sys.exit("Пароли не совпали.")
    return p1


def project_cmd(args):
    """Проект = комплект документов со своим глоссарием заказчика. Новая версия глоссария существующего проекта
    загружается экспертом в веб-интерфейсе (Эксперт → Импорт глоссария заказчика)."""
    from server import config, glossary_db as GD
    if args.cmd == "list":
        for pj in sorted(config.PROJECTS.glob("*/project.json")):
            print(pj.parent.name)
        return
    if not re.fullmatch(r"[\w-]{1,40}", args.name):
        sys.exit("Имя проекта — латиница, цифры, «-» и «_», до 40 знаков.")
    d = config.PROJECTS / args.name
    if (d / "project.json").exists():
        sys.exit(f"Проект {args.name} уже есть. Новую версию глоссария загружает эксперт: "
                 "Эксперт → Импорт глоссария заказчика.")
    import zipfile
    from docx.opc.exceptions import PackageNotFoundError
    try:
        rules, entries = GD.parse_upload(args.glossary)
    except (zipfile.BadZipFile, PackageNotFoundError, KeyError):
        sys.exit("Глоссарий не прочитан: файл повреждён или это не docx/xlsx. Нужен файл md, docx или xlsx "
                 "с таблицами «венгерский | русский | комментарий».")
    except (ValueError, OSError) as e:
        sys.exit(f"Глоссарий не прочитан: {e}")
    if not entries:
        sys.exit("В файле не нашлось пар «венгерский — русский». Нужны таблицы: венгерский | русский | комментарий.")
    (d / "data").mkdir(parents=True, exist_ok=True)
    (d / "data" / "glossary.md").write_text(GD.entries_md(f"Глоссарий {args.name}", rules, entries), encoding="utf-8")
    (d / "project.json").write_text(json.dumps({"glossary": "data/glossary.md", "tm": "data/tm.json"}, ensure_ascii=False,
                                               indent=1), encoding="utf-8")
    print(f"Проект {args.name}: терминов {len(entries)}{', правила применения есть' if rules else ''}. "
          "В веб-интерфейсе он уже в списке проектов.")


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    ap = argparse.ArgumentParser(description="Команды администратора переводчика")
    sub = ap.add_subparsers(dest="what", required=True)
    u = sub.add_parser("user").add_subparsers(dest="cmd", required=True)
    a = u.add_parser("add"); a.add_argument("login"); a.add_argument("--name", default="")
    a.add_argument("--role", choices=ROLES, default="user"); a.add_argument("--password")
    r = u.add_parser("role"); r.add_argument("login"); r.add_argument("role", choices=ROLES)
    p = u.add_parser("passwd"); p.add_argument("login"); p.add_argument("--password")
    for name in ("disable", "enable"):
        u.add_parser(name).add_argument("login")
    u.add_parser("list")
    sub.add_parser("backup")
    pr = sub.add_parser("project").add_subparsers(dest="cmd", required=True)
    pa = pr.add_parser("add"); pa.add_argument("name"); pa.add_argument("glossary")
    pr.add_parser("list")
    args = ap.parse_args()
    db.init()
    if args.what == "backup":
        from server import maintenance
        print(f"Копия базы: {maintenance.backup()}")
        return
    if args.what == "project":
        project_cmd(args)
        return

    if args.cmd == "list":
        for row in db.q("select login, name, role, active, token_enc is not null as tok from users order by login"):
            print(f"{row['login']:<16} {row['role']:<7} {'активен' if row['active'] else 'отключён':<9} "
                  f"{'токен есть' if row['tok'] else 'токена нет':<11} {row['name']}")
        return
    if args.cmd != "add" and not db.one("select 1 from users where login=?", (args.login,)):
        sys.exit(f"Нет пользователя {args.login}.")
    if args.cmd == "add":
        if db.one("select 1 from users where login=?", (args.login,)):
            sys.exit(f"Пользователь {args.login} уже есть.")
        pw = ask_password(args.password)
        if len(pw) < 8:
            sys.exit("Пароль короче 8 знаков.")
        db.x("insert into users(login, name, role, pw_hash, created) values (?, ?, ?, ?, ?)",
             (args.login, args.name, args.role, security.hash_pw(pw), time.time()))
        print(f"Создан {args.login} ({args.role}).")
    elif args.cmd == "role":
        db.x("update users set role=? where login=?", (args.role, args.login))
        print(f"{args.login}: роль {args.role}.")
    elif args.cmd == "passwd":
        pw = ask_password(args.password)
        if len(pw) < 8:
            sys.exit("Пароль короче 8 знаков.")
        db.x("update users set pw_hash=? where login=?", (security.hash_pw(pw), args.login))
        db.x("delete from sessions where user_id=(select id from users where login=?)", (args.login,))
        print(f"{args.login}: пароль изменён, все входы сброшены.")
    else:
        active = int(args.cmd == "enable")
        db.x("update users set active=? where login=?", (active, args.login))
        if not active:
            db.x("delete from sessions where user_id=(select id from users where login=?)", (args.login,))
        print(f"{args.login}: {'включён' if active else 'отключён'}.")


if __name__ == "__main__":
    main()
