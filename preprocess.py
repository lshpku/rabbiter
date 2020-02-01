import math

with open('schools.txt', 'r') as f:
    schools = f.readlines()
    schools = [''.join(i.split()) for i in schools]

with open('routes.txt', 'r') as f:
    routes = []
    route = None
    for i in f.readlines():
        i = ''.join(i.split())
        if not route:
            route = i
            continue
        if not i:
            route = i
            continue
        routes.append((i, route))


def simularity(a, b):
    a = set(a)
    b = set(b)
    return len(a & b) / math.sqrt(max(len(a), len(b)))


for j, k in routes:
    rank = [(-simularity(j, i), i) for i in schools]
    rank.sort()
    i = '\t'.join([r[1] for r in rank[:5]])
    print('{}\t{}\t{}'.format(k, j, i))

with open('sexp.txt', 'r') as f:
    sexp = f.readlines()
    sexp = [''.join(i.split()) for i in sexp]


print(len(sexp))
if len(sexp) != len(set(sexp)):
    print('fuck!')
else:
    print('ok')
for i in schools:
    if i not in sexp:
        print('excluded:', i)
