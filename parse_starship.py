"""Convert a starship data file in VB6 format to JSON."""


import json
import struct
import sys


def read_unpack(binary_file, fmt):
    """Read bytes from a binary file sized according to the given struct format."""
    return struct.unpack(fmt, binary_file.read(struct.calcsize(fmt)))


def get_ship_type(type_code):
    types = ['Short Range Patrol Craft',
             'Civilian Transport/Freighter',
             'Civilian Passenger Liner',
             'Military Transport/Freighter',
             'Military Troop Transport',
             'Military Combat Starship']
    return types[type_code]


def get_hyperdrive(hd_type):
    hdrives = ['None', 'Standard', 'Improved']
    return hdrives[hd_type]


def get_speed(val):
    if val >= 4:
        return str(val - 3)
    if val == 1:
        return '1/8'
    elif val == 2:
        return '1/4'
    elif val == 3:
        return '1/2'
    elif val == 0:
        return '0'
    else:
        raise RuntimeError


def get_techbase(tb):
    strs = ['Common', 'New Republic', 'Imperial', 'Herald', 'Ploxus']
    return strs[tb]


def parse_starship(filename):
    ship = {}
    with open(filename, 'rb') as f:
        ship_record = read_unpack(f, '<35sd35s2hd6h')
        ship['name'] = ship_record[0].decode().strip()
        ship['manufacturer'] = ship_record[2].decode().strip()
        ship['type'] = get_ship_type(ship_record[3])
        ship['mass_tons'] = '{:.2f}'.format(ship_record[1])
        ship['length_m'] = ship_record[4]
        ship['cargo_tons'] = ship_record[5]
        ship['shields'] = ship_record[8]
        ship['hull'] = ship_record[6]
        ship['atmosphere_capable'] = ship_record[7]
        ship['speed'] = get_speed(ship_record[9])
        ship['hyperdrive'] = get_hyperdrive(ship_record[10])
        ship['techbase'] = get_techbase(ship_record[11])
    return ship


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print('usage:', 'python', sys.argv[0], '<sw2 file>')
        sys.exit(1)

    print(json.dumps(parse_starship(sys.argv[1]), indent=4))
