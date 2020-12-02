# Modified.
# Original script source:
# http://blog.marcbelmont.com/2012/10/script-to-extract-email-attachments.html

# Usage:
# Run the script from a folder with file "all.mbox"
# Attachments will be extracted into subfolder "attachments" 
# with prefix "n " where "n" is an order of attachment in mbox file. 

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
            name = attachment_number
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
                    name = attachment_number
                    print('Could not decode %s %s attachment name.' % (mid, name))

        name = '%s %s' % (mid, name)

        try:
            with open(save_to + name, 'wb') as f:
                f.write(part.get_payload(decode=True))
        except OSError as e:
            if e.errno == errno.ENAMETOOLONG:

                extension = pathlib.Path(name).suffix
                extension = extension if len(extension) <= 20 else ''

                short_name = '%s %s%s' % (mid, attachment_number, extension)

                with open(save_to + short_name, 'wb') as f:
                    f.write(part.get_payload(decode=True))
            else:
                raise

    except:
        traceback.print_exc()
        global failed
        failed = failed + 1


def check_part(mid, part, attachments_counter):
    mime_type = part.get_content_type()
    if (part.get_content_disposition() == 'attachment') \
            or ((part.get_content_disposition() != 'inline') and (part.get_filename() is not None)):
        save(mid, part, attachments_counter)
    elif (mime_type.startswith('application/') and not mime_type == 'application/javascript') \
            or mime_type.startswith('model/') \
            or mime_type.startswith('audio/') \
            or mime_type.startswith('video/'):
        if part.get_content_disposition() == 'inline':
            print('Extracting inline part...')
        print('Ignoring Content-disposition... Message id = %s, Content-type = %s.' % (mid, mime_type))
        save(mid, part, attachments_counter)
    elif prefs['extract_inline_images'] and mime_type.startswith('image/'):
        save(mid, part, attachments_counter, True)
    elif part.is_multipart():
        for p in part.get_payload():
            check_part(mid, p, attachments_counter)


def process_message(mid):
    msg = mb.get_message(mid)
    if msg.is_multipart():
        attachments_counter = {'value': 0, 'inline_image': 0}
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
print('Total:  %s' % (total))
print('Failed: %s' % (failed))
