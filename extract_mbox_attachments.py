#!/usr/bin/env python3
# -*- coding: utf-8 -*-

# Modified.
# Original script source:
# http://blog.marcbelmont.com/2012/10/script-to-extract-email-attachments.html
# https://web.archive.org/web/20150312172727/http://blog.marcbelmont.com/2012/10/script-to-extract-email-attachments.html

# Usage:
# Run the script from a folder with file "all.mbox"
# Attachments will be extracted into subfolder "attachments" 
# with prefix "m " where m is a message ID in mbox file.

# ---------------
# Please check the unpacked files
# with an antivirus before opening them!

# ---------------
# I make no representations or warranties of any kind concerning
# the software, express, implied, statutory or otherwise,
# including without limitation warranties of title, merchantability,
# fitness for a particular purpose, non infringement, or the
# absence of latent or other defects, accuracy, or the present or
# absence of errors, whether or not discoverable, all to the
# greatest extent permissible under applicable law.

import errno
import mailbox
import os
import pathlib  # since Python 3.4
import re
import traceback
from email.header import decode_header

prefs = {
    'file': 'all.mbox',
    'save_to': 'attachments/',
    'extract_inline_images': True,
    'start': 0,
    'stop': 100000000000  # On which message to stop (not included).
}

assert os.path.isfile(prefs['file'])

mb = mailbox.mbox(prefs['file'])

if not os.path.exists(prefs['save_to']):
    os.makedirs(prefs['save_to'])

inline_image_folder = os.path.join(prefs['save_to'], 'inline_images/')

if prefs['extract_inline_images'] and not os.path.exists(inline_image_folder):
    os.makedirs(inline_image_folder)

total = 0
failed = 0


def to_file_path(save_to, name):
    return os.path.join(save_to, name)


def get_extension(name):
    extension = pathlib.Path(name).suffix
    return extension if len(extension) <= 20 else ''


def resolve_name_conflicts(save_to, name, file_paths, attachment_number):
    file_path = to_file_path(save_to, name)

    START = 1
    iteration_number = START

    while os.path.normcase(file_path) in file_paths:
        extension = get_extension(name)
        iteration = '' if iteration_number <= START else ' (%s)' % iteration_number
        new_name = '%s attachment %s%s%s' % (name, attachment_number, iteration, extension)
        file_path = to_file_path(save_to, new_name)
        iteration_number += 1

    file_paths.append(os.path.normcase(file_path))
    return file_path


# Whitespaces: tab, carriage return, newline, vertical tab, form feed.
FORBIDDEN_WHITESPACE_IN_FILENAMES = re.compile('[\t\r\n\v\f]+')
OTHER_FORBIDDEN_FN_CHARACTERS = re.compile('[/\\\\\\?%\\*:\\|"<>\0]')


def filter_fn_characters(s):
    result = re.sub(FORBIDDEN_WHITESPACE_IN_FILENAMES, ' ', s)
    result = re.sub(OTHER_FORBIDDEN_FN_CHARACTERS, '_', result)
    return result


def save(mid, part, attachments_counter, inline_image=False):
    global total
    total = total + 1

    try:
        if inline_image:
            attachments_counter['inline_image'] += 1
            attachment_number = 'ii' + str(attachments_counter['inline_image'])
            save_to = inline_image_folder
        else:
            attachments_counter['value'] += 1
            attachment_number = attachments_counter['value']
            save_to = prefs['save_to']

        if part.get_filename() is None:
            name = str(attachment_number)
            print('Filename is none: %s %s.' % (mid, name))
        else:
            decoded_name = decode_header(part.get_filename())

            if isinstance(decoded_name[0][0], str):
                name = decoded_name[0][0]
            else:
                try:
                    name_encoding = decoded_name[0][1]
                    name = decoded_name[0][0].decode(name_encoding)
                except:
                    name = str(attachment_number)
                    print('Could not decode %s %s attachment name.' % (mid, name))

        name = filter_fn_characters(name)
        name = '%s %s' % (mid, name)

        previous_file_paths = attachments_counter['file_paths']

        try:
            fn = resolve_name_conflicts(save_to, name,
                                        previous_file_paths,
                                        attachment_number)
            with open(fn, 'wb') as f:
                f.write(part.get_payload(decode=True))
        except OSError as e:
            if e.errno == errno.ENAMETOOLONG:

                extension = get_extension(name)
                short_name = '%s %s%s' % (mid, attachment_number, extension)

                fn = resolve_name_conflicts(save_to, short_name,
                                            previous_file_paths,
                                            attachment_number)
                with open(fn, 'wb') as f:
                    f.write(part.get_payload(decode=True))
            else:
                raise

    except:
        traceback.print_exc()
        global failed
        failed = failed + 1


def check_part(mid, part, attachments_counter):
    mime_type = part.get_content_type()
    if part.is_multipart():
        for p in part.get_payload():
            check_part(mid, p, attachments_counter)
    elif (part.get_content_disposition() == 'attachment') \
            or ((part.get_content_disposition() != 'inline') and (part.get_filename() is not None)):
        save(mid, part, attachments_counter)
    elif (mime_type.startswith('application/') and not mime_type == 'application/javascript') \
            or mime_type.startswith('model/') \
            or mime_type.startswith('audio/') \
            or mime_type.startswith('video/'):
        message_id_content_type = 'Message id = %s, Content-type = %s.' % (mid, mime_type)
        if part.get_content_disposition() == 'inline':
            print('Extracting inline part... ' + message_id_content_type)
        else:
            print('Other Content-disposition... ' + message_id_content_type)
        save(mid, part, attachments_counter)
    elif prefs['extract_inline_images'] and mime_type.startswith('image/'):
        save(mid, part, attachments_counter, True)


def process_message(mid):
    msg = mb.get_message(mid)
    if msg.is_multipart():
        attachments_counter = {
            'value': 0,
            'inline_image': 0,
            'file_paths': []
        }
        for part in msg.get_payload():
            check_part(mid, part, attachments_counter)


for i in range(prefs['start'], prefs['stop']):
    try:
        process_message(i)
    except KeyError:
        print('The whole mbox file was processed.')
        break
    if i % 1000 == 0:
        print('Messages processed: {}'.format(i))

print()
print('Total:  %s' % total)
print('Failed: %s' % failed)
