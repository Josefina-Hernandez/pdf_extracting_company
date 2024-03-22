import re

def if_email(string):
    pattern = re.compile(r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$')
    if pattern.match(string):
        return True
    else:
        return False

a=if_email(string='GTC.5555@@@tom.com')

print(a)