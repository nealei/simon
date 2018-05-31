#! /usr/bin/env python
""" Use existing road file to create new overtaking lane option file(s).

By Neale Irons version 10/05/2018 (CC BY-SA 4.0)
Known bugs:
"""

from __future__ import print_function
import argparse
import hashlib
import os
import shutil
from shutil import move
import sys
import tempfile
from tempfile import NamedTemporaryFile
import logging


def main():

    def usage():
        parser = argparse.ArgumentParser(description=
                "Use existing road files to create new overtaking lane.",
                epilog=
                ' eg: OptionEditor.py Demo.ROD 1.5 2.2 P Option.ROD'
                ' "Option description"')
        parser.add_argument("Infile", help="Input ROD file")
        parser.add_argument("Start", type=float, help="Start km")
        parser.add_argument("End", type=float, help="End km")
        parser.add_argument("Direction",
                            help="Direction P, C or B (direction 1 and/or 2)",
                            choices=["P", "C", "B"])
        parser.add_argument("Outfile", help="Output ROD file")
        parser.add_argument('"Description"', nargs='?',
                            help="Description line for Outfile")
        args = parser.parse_args()

    def validate_command_line(argv, otlane_direction_codes):
        """ Check command line is valid. """
        if len(argv) < 6 or len(argv) > 7:
            usage()
            sys.exit()

        src = argv[1]
        otl_start_km = round(float(argv[2]), 1)
        otl_end_km = round(float(argv[3]), 1)
        otl_direction = argv[4]
        dst = argv[5]
        if len(argv) == 7:
            description = argv[6]
            description = description[:50]
        else:
            description = ''

        if otl_start_km > otl_end_km:
            sys.exit("Invalid start/end overtaking lane chainage - "
                     "start must be less than end. Exiting...")

        if otl_direction not in OTLANE_DIRECTION_CODES:
            sys.exit("Direction %s is invalid, must be P, C or B. "
                     "Exiting..." % argv[4])

        logging.debug('Infile Start End Dir Outfile: %s, %3.1f, %3.1f, %s, %s',
                      src, otl_start_km, otl_end_km, otl_direction, dst)

        return(src, otl_start_km, otl_end_km, otl_direction, dst, description)

    def validate_file(src, dst, hashfile, record_size):
        """ File integrity checks. """
        if not os.path.isfile(src):
            sys.exit("File path %s does not exist. Exiting..." % src)

        src_name, src_ext = os.path.splitext(src)
        if src != dst:
            # Validate support files
            src_name, src_ext = os.path.splitext(src)
            if not os.path.isfile(src_name + '.MLT'):
                sys.exit("Requires support file %s. Exiting..."
                         % (src_name + '.MLT'))
            if not os.path.isfile(src_name + '.OBS'):
                sys.exit("Requires support file %s. Exiting..."
                         % (src_name + '.OBS'))

        # Validate HashCode
        if not hashcode_register(HASHFILE, src, "validate"):
            sys.exit("Invalid file - %s failed HashCode validation. Exiting..."
                     % src)

        # Read header
        with open(src) as fp:
            # Fails - "Invalid file format - first 9 lines must be file header
            for x in range(1, 10):
                line = fp.readline()
            header_length = fp.tell()       # Calculate header length
            fp.seek(0, 2)
            file_length = fp.tell()         # Get file length

        if (file_length - header_length) % RECORD_SIZE != 0:
            sys.exit("Invalid file size - must be 9 line header and records"
                     " must be %s characters long. Exiting..." % RECORD_SIZE)

        nrecs = (file_length - header_length) // RECORD_SIZE
        if nrecs < 1:
            sys.exit("Invalid file format"
                     " - file must contain at least one record. Exiting...")

        logging.debug('src_name header_length nrecs: %s, %i, %i',
                      src_name, header_length, nrecs)

        return(src_name, header_length, nrecs)

    def hashcode_register(hashfile, src, service):
        """Validate or update register hashcode returning true/false.

        By Neale Irons version 04/03/2018 (CC BY-SA 4.0)

        Keyword arguments:
        hashfile - hashcode_register filename
        src - filename to service
        service - using src, validate or update register

        Returns True for success or False for failure

        import hashlib
        import os
        """
        return True  # Works but disabled, future feature
        # hashcode_register must exist or error
        if not os.path.isfile(hashfile):
            return False
        hashcode = hashlib.md5(open(src, 'rb').read()).hexdigest()
        # hashcode is lowercase

        if service == 'validate':
            with open(hashfile) as hf:
                for line in hf:
                    if hashcode in line.lower():
                        # hashcode found but filename might not match
                        return True

        elif service == "update":
            with open(hashfile) as hf:
                for lcnt, line in enumerate(hf):
                    register_hash, register_file = line[:-1].split(" *")
                    if register_file == src:
                        if hashcode == register_hash.lower():
                            # File already registered
                            return True
                        else:
                            # Remove invalid hashcode record from register
                            hf.seek(0)
                            with tempfile.NamedTemporaryFile(delete=False) as tmp:
                                for lnum, line in enumerate(hf):
                                    if lnum != lcnt:
                                        tmp.write(line)
                            # Rename tmp as file
                            hf.close()
                            os.unlink(hashfile)
                            os.rename(tmp.name, hashfile)
                            break
            with open(hashfile, 'a') as hf:
                # Append new hashcode file record to register
                hf.write(hashcode + ' *' + src + '\n')
                return True

        return False

    def edit_file(src, dst, header_length, otl_start_km, otl_end_km,
                  otl_direction, odo_step_km, record_size, barrier_column1,
                  barrier_column2, barrier_value, otlane_column1,
                  otlane_column2, otlane_value, hashfile, nrecs, description):
        """ Change file records. """
        # Make temporary copy of source file to change
        # dst_tmp = tempfile.NamedTemporaryFile(delete=False) # Not working **
        dst_tmp = src + ".tmp"                                # Workaround ***
        shutil.copy2(src, dst_tmp)

        # Open file for random read/write
        with open(dst_tmp, 'r+b') as fp:
            fp.seek(header_length, 0)
            odo_start_km = float((fp.read(9)))
            fp.seek(RECORD_SIZE * -1, 2)
            odo_end_km = float((fp.read(9)))

            if otl_start_km < odo_start_km:
                logging.error('otl_start_km, odo_start_km:   %f, %f',
                              otl_start_km, odo_start_km)
                sys.exit("Invalid start overtaking lane chainage -"
                         " Overtaking lane can't start before first road"
                         " chainage. Exiting...")
            if otl_end_km > odo_end_km + ODO_STEP_KM:
                logging.error('otl_end_km, odo_end_km:       %f, %f',
                              otl_start_km, odo_start_km)
                sys.exit("Invalid end overtaking lane chainage - Overtaking "
                         "lane can't end after last road chainage. Exiting...")

            if int(round((odo_end_km - odo_start_km)/ODO_STEP_KM) + 1) != nrecs:
                sys.exit("Invalid file format - wrong number of lines "
                         "between start and end distances. Exiting...")

            logging.debug('odo_start_km  odo_end_km:     %3.3f, %3.3f',
                          odo_start_km, odo_end_km)

            while otl_start_km + .000001 < otl_end_km:   # 1mm rounding error
                nrec = (otl_start_km - .000001 - odo_start_km) // \
                        ODO_STEP_KM + 1
                fpos = header_length + (nrec * RECORD_SIZE) - 1

                logging.debug('nrec fpos  otl_start_km: %3.0f, %i, %3.1f',
                              nrec, fpos, otl_start_km)

                fp.seek(BARRIER_COLUMN1 + fpos)
                fp.write(BARRIER_VALUE)
                fp.seek(BARRIER_COLUMN2 + fpos)
                fp.write(BARRIER_VALUE)

                if otl_direction == "P" or otl_direction == "B":
                    fp.seek(OTLANE_COLUMN1 + fpos)
                    fp.write(OTLANE_VALUE)
                if otl_direction == "C" or otl_direction == "B":
                    fp.seek(OTLANE_COLUMN2 + fpos)
                    fp.write(OTLANE_VALUE)
                otl_start_km += ODO_STEP_KM

        if dst != src:
            # Create new supporting files
            dst_name, dst_ext = os.path.splitext(dst)
            shutil.copy2(src_name + '.MLT', dst_name + '.MLT')
            replace_textfile_line(dst_name + '.MLT', dst_name + '.MLT', '')
            shutil.copy2(src_name + '.OBS', dst_name + '.OBS')

        replace_textfile_line(dst_tmp, dst, description)
        move(dst_tmp, dst)
        hashcode_register(HASHFILE, dst, "update")

    def replace_textfile_line(text_filename, new_text, description):
        """ Replace first line with new_text.

        Adapted from ReplFHdr.py
        By Neale Irons version 14/02/2018 (CC BY-SA 4.0)
        import sys, os, tempfile
        """

        with open(text_filename) as src:
            with tempfile.NamedTemporaryFile('w',
              dir=os.path.dirname(text_filename), delete=False) as dst:
                line = src.readline()               # Read first line
                dst.write(new_text + '\n')          # Replace first line
                if description:                     # If description not empty
                    line = src.readline()           # Read second line
                    dst.write(description + '\n')   # Replace second line
                shutil.copyfileobj(src, dst)        # Copy rest of file
        os.unlink(text_filename)                    # remove old version
        os.rename(dst.name, text_filename)          # rename new version

    # Constants
    ODO_STEP_KM = 0.1
    RECORD_SIZE = 109
    BARRIER_COLUMN1 = 11
    BARRIER_COLUMN2 = 16
    BARRIER_VALUE = '-'     # Debug - use '+' instead of '-'
    OTLANE_COLUMN1 = 22
    OTLANE_COLUMN2 = 27
    OTLANE_VALUE = 'T'      # Debug - use 'Y' instead of 'T'
    OTLANE_DIRECTION_CODES = "PCB"
    HASHFILE = "Hashcode.MD5"

    # Enable logging with level=logging.DEBUG, disable with .ERROR
    logging.basicConfig(level=logging.ERROR,
                        format='%(levelname)s - %(message)s')

    # Validate command line
    src, otl_start_km, otl_end_km, otl_direction, dst, description = validate_command_line(
        sys.argv, OTLANE_DIRECTION_CODES)

    # Validate file
    src_name, header_length, nrecs = validate_file(src, dst, HASHFILE,
                                                   RECORD_SIZE)

    # Edit file
    edit_file(src, dst, header_length, otl_start_km, otl_end_km, otl_direction,
              ODO_STEP_KM, RECORD_SIZE, BARRIER_COLUMN1, BARRIER_COLUMN2,
              BARRIER_VALUE, OTLANE_COLUMN1, OTLANE_COLUMN2, OTLANE_VALUE,
              HASHFILE, nrecs, description)

    print(" done.")


if __name__ == '__main__':
    main()
