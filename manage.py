# -*- coding: utf-8 -*-
"""Команды администратора веб-сервиса (в контейнере: docker compose exec hut python manage.py …).

  python manage.py user add <логин> [--name "Иванов М."] [--role user|expert|admin] [--password …]
  python manage.py user role <логин> <роль>
  python manage.py user passwd <логин> [--password …]
  python manage.py user disable <логин> | enable <логин>
  python manage.py user list
  python manage.py backup                       — копия базы сейчас (сервис делает её и сам раз в сутки)
Без --password пароль спрашивается с клавиатуры (так он не остаётся в истории команд)."""
import sys
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
    args = ap.parse_args()
    db.init()
    if args.what == "backup":
        from server import maintenance
        print(f"Копия базы: {maintenance.backup()}")
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
