#!/usr/bin/env python3
# encoding: UTF-8

from __future__ import print_function
import os


def pip_install(package):
    """
    Install the package using pip
    """
    try:
        import pip
        pip.main(["install", "--upgrade", package, "--user"])
    except Exception as e:
        print("Unable to install %s using pip." % package)
        print("Exception:", e)
        exit(-1)


try:
    from ldap3 import Server, Connection, ALL
except ImportError:
    pip_install("ldap3")
    print("Software installed, restart program. Exiting in 5 seconds.")
    sleep(5)
    exit(0)

# Global variables
con = Connection("ldap.uio.no", auto_bind=True)
base = "cn=people,dc=uio,dc=no"


def get_input(prompt):
    "Get a nonempty input from the user"
    read = ""
    while not read:
        read = input(prompt).strip()
    return read


def print_file(filename):
    "Send a text file to the default Windows printer"
    import subprocess
    try:
        subprocess.call(["notepad", "/p", filename])
#         subprocess.call(["notepad", filename])
    except Exception as e:
        print("Unable to start Notepad:", e)


def get_user_ou(user):
    "Get the printable organisation unit of user"
    return format_ou(user.eduPersonPrimaryOrgUnitDN)


def xstr(item):
    """
    Get string representation, or empty string if item is None
    Example:
    >>> xstr(None)
    ''
    >>> xstr("test")
    'test'
    """
    if item:
        return str(item)
    else:
        return ""


def make_criteria(string):
    """
    Make conjunctive search criteria from a name string

    Example:
    >>> make_criteria("John ")
    '(&(cn=*John*))'
    >>> make_criteria(" John Doe Henry ")
    '(&(cn=*John*)(cn=*Doe*)(cn=*Henry*))'
    """
    terms = ["(cn=*%s*)" % name for name in string.split()if name]
    return "(&%s)" % "".join(terms)


def format_ou(ou_string):
    """
    Make a human-readable string of the OU

    Example:
    >>> format_ou("ou=UJUR,ou=UB,ou=UIO,cn=organization,dc=uio,dc=no")
    'UJUR/UB/UIO'
    """
    parts = [item[3:] for item in xstr(ou_string).split(",") if item.startswith("ou=")]
    return "/".join(parts)


def print_person(entry):
    "Print address slip for an LDAP entry"
    filename = "address-temp.txt"
    try:
        with open(filename, "w") as out:
            print(entry.cn, get_user_ou(entry), xstr(entry.street).replace("$", "\n"), sep="\n", end="\n", file=out, flush=True)
        print_file(filename)
        os.remove(filename)
    except Exception as e:
        print(e)


def find_person():
    "Lookup a user in LDAP"
    name = get_input("Name: ").strip().lower()
    if name in["quit", "exit"]:
        exit(0)

    query = make_criteria(name)
#     print(query)
    con.search(base, query, attributes=["cn", "postalAddress", "eduPersonPrimaryOrgUnitDN", "street", "uioShortPhone"])
#     print(con.entries)

    for index, user in enumerate(con.entries):
        print("[%2d]  %-40s  %-30s  %s" % (index, user.cn, get_user_ou(user), xstr(user.uioShortPhone)))

    if len(con.entries) == 0:
        print("No match for %s" % name)
        print("Names with more than 50 matches cannot be displayed")
        return

    choice = ""
    while not (isinstance(choice, int) and choice >= 0 and choice < len(con.entries)):
        choice = get_input("Select person 0-%d (a aborts): " % (len(con.entries) - 1))
        if choice == "a":
            return
        try:
            choice = int(choice)
        except ValueError:
            pass

    entry = con.entries[choice]
    print_person(entry)


if __name__ == '__main__':
    while True:
        try:
            find_person()
        except KeyboardInterrupt:
            print("Bye")
            exit(0)
