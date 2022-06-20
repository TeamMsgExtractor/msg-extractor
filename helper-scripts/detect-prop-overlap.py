#! /usr/bin/env python3
import glob
import re
import sys


def main(args):
    """
    Checks the specified file(s) for overlapping properties. Properties are
    overlapping if they share the same variable name or come from the same
    property. This may be intentional for some properties.
    """
    if len(sys.argv) < 2:
        print('Please specify a file to read.')
        sys.exit(1)

    pattern = re.compile(r"(?<=self._ensureSet)((Named)|(Property)|(Typed))?\('(.*?)', '(.*?)'")

    for patt in sys.argv[1:]:
        for name in glob.glob(patt):
            with open(name, 'r', encoding = 'utf-8') as f:
                data = f.read()

            names = [x.group(5) for x in pattern.finditer(data)]
            ids = [x.group(6) for x in pattern.finditer(data)]

            duplicateNamesFound = len(names) != len(list(set(names)))
            duplicateIdsFound = len(ids) != len(list(set(names)))

            print(name)

            if duplicateNamesFound:
                print('\tVariable Names:')
                counts = {x: 0 for x in names}
                for x in names:
                    counts[x] += 1
                for x in counts:
                    if counts[x] > 1:
                        print(f'\t\t{x}')

            if duplicateIdsFound:
                print('\tIDs:')
                counts = {x: 0 for x in ids}
                for x in ids:
                    counts[x] += 1
                for x in counts:
                    if counts[x] > 1:
                        print(f'\t\t{x}')

            if not duplicateIdsFound and not duplicateNamesFound:
                print('\tNo Duplicates detected.')

            print()


if __name__ == '__main__':
    main(sys.argv)
