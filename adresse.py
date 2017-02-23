from __future__ import print_function
from ldap3 import Server, Connection, ALL
import os

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


def print_person(entry):
    "Print address slip for an LDAP entry"
    filename = "address-temp.txt"
    try:
        with open(filename, "w") as out:
            print(entry.cn, str(entry.postalAddress).replace("$", "\n"), sep="\n", end="\n", file=out, flush=True)
        print_file(filename)
        os.remove(filename)
    except Exception as e:
        print(e)


def get_user():
    "Lookup a user in LDAP"
    name = get_input("Name: ")
    search_string = "(cn =* %s)" % name
    con.search(base, search_string, attributes=["cn", "postalAddress"])

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
