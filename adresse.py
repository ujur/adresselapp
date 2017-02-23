from __future__ import print_function
from ldap3 import Server, Connection, ALL
import os

# Global variables
con = Connection("ldap.uio.no", auto_bind=True)
base = "cn=people,dc=uio,dc=no"


def print_file(filename):
    "Send a text file to the default Windows printer"
    import subprocess
    subprocess.call(["notepad", "/p", filename])
#     subprocess.call(["notepad", filename])


def get_user():
    "Lookup a user in LDAP"
    name = input("Name: ")
    search_string = "(cn =* %s)" % name
    con.search(base, search_string, attributes=["cn", "postalAddress"])

    for index, user in enumerate(con.entries):
        print("[%d]" % index, user.cn)

    if len(con.entries) == 0:
        print("No match for %s" % name)
        print("Names with more than 200 matches cannot be displayed")
        return

    choice = input("Select person (a aborts): ")
    if choice == "a":
        return

    entry = con.entries[int(choice)]

    filename = "address-temp.txt"

    try:
        with open(filename, "w") as out:
            print(entry.cn, str(entry.postalAddress).replace("$", "\n"), sep="\n", end="\n", file=out, flush=True)

        print_file(filename)
        os.remove(filename)
    except Exception as e:
        print(e)


if __name__ == '__main__':
    get_user()
