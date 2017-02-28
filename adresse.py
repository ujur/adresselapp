#!/usr/bin/env python3

from __future__ import print_function
import os


def pip_install(package):
    """
    Install the package using pip
    """
    import pip
    try:
        pip.main(["install", "--upgrade", package, "--user"])
    except:
        print("Unable to install %s using pip." % package)
#         print("Error: %s: %s" % (exc_info()[0], exc_info()[1]))
        exit


try:
    from ldap3 import Server, Connection, ALL
except ImportError:
    pip_install("ldap3")
    from ldap3 import Server, Connection, ALL

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
    subprocess.call(["notepad", "/p", filename])
#     subprocess.call(["notepad", filename])


def make_criteria(string):
    """
    Make conjunctive search criteria from a name string

    Example:
    >>> make_criteria("John Doe")
    '(&(cn=*John*)(cn=*Doe*))'
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
    parts = [item[3:] for item in str(ou_string).split(",") if item.startswith("ou=")]
    return "/".join(parts)


def print_person(entry):
    "Print address slip for an LDAP entry"
    filename = "address-temp.txt"
    try:
        with open(filename, "w") as out:
            print(entry.cn, format_ou(entry.eduPersonPrimaryOrgUnitDN), str(entry.street).replace("$", "\n"), sep="\n", end="\n", file=out, flush=True)
#             print(entry.cn, format_ou(entry.eduPersonPrimaryOrgUnitDN), str(entry.street).replace("$", "\n"), sep="\n", end="\n", flush=True)
        print_file(filename)
        os.remove(filename)
    except Exception as e:
        print(e)


def get_user():
    "Lookup a user in LDAP"
    name = get_input("Name: ")
    query = make_criteria(name)
#     print(query)
    con.search(base, query, attributes=["cn", "postalAddress", "eduPersonPrimaryOrgUnitDN", "street"])
#     print(con.entries)

    for index, user in enumerate(con.entries):
        print("[%d]" % index, user.cn)

    if len(con.entries) == 0:
        print("No match for %s" % name)
        print("Names with more than 200 matches cannot be displayed")
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
    get_user()
