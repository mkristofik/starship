"""Convert the starship weapons data file in VB6 format to JSON."""

import json
import struct


def get_techbase(tb):
    strs = ['Common', 'New Republic', 'Imperial', 'Herald', 'Ploxus']
    return strs[tb]


def parse_weapons():
    with open('weapons.db', 'rb') as f:
        count = 0
        for record in struct.iter_unpack('<25s6sdd15s6sh', f.read()):
            count += 1
            yield {'id': count,
                'name': record[0].decode().strip(),
                'damage': record[1].decode().strip(),
                'mass': record[2],
                'power': record[3],
                'range': record[4].decode().strip(),
                'tohit': record[5].decode().strip(),
                'techbase': get_techbase(record[6])
            }


if __name__ == '__main__':
    print(json.dumps(list(parse_weapons()), indent=4))
