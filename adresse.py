from ldap3 import Server, Connection, ALL

con = Connection("ldap.uio.no", auto_bind=True)
# print(con)

base = "cn=people,dc=uio,dc=no"

con.search(base, "(cn = *Winge)")
print(con.entries)