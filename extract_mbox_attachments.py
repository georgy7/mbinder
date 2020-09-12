# Modified.
# Original script source:
# http://blog.marcbelmont.com/2012/10/script-to-extract-email-attachments.html

# Usage:
# Run the script from a folder with file "all.mbox"
# Attachments will be extracted into subfolder "attachments" 
# with prefix "n " where "n" is an order of attachment in mbox file. 

import mailbox, pickle, traceback, os
from email.header import decode_header

mb = mailbox.mbox('all.mbox')

prefs_path = '.save-attachments'
save_to = 'attachments/'

if not os.path.exists(save_to): os.makedirs(save_to)

prefs = {
    'start': 0,
    'stop': 100000000000   # On which message to stop (not included).
}


total = 0
failed = 0

def save_attachments(mid):
    msg = mb.get_message(mid)
    if msg.is_multipart():
        for part in msg.get_payload():
            if part.get_content_type() != 'application/octet-stream':
                continue
            
            global total
            total = total + 1

            print()
            try:
                decoded_name = decode_header(part.get_filename())
                print(decoded_name)
                
                if isinstance(decoded_name[0][0], str):
                    name = decoded_name[0][0]
                else:
                    name_encoding = decoded_name[0][1]
                    name = decoded_name[0][0].decode(name_encoding)
                
                name = '%s %s' % (total, name)
                # print('Saving %s' % (name))
                with open(save_to + name, 'wb') as f:
                    f.write(part.get_payload(decode=True))
            except:
                traceback.print_exc()
                global failed
                failed = failed + 1


for i in range(prefs['start'], prefs['stop']):
    try:
        save_attachments(i)
    except KeyError:
        print('The whole mbox file was processed.')
        break
    if i % 1000 == 0:
        print('Messages processed: {}'.format(i))

print()
print('Total:  %s' % (total))
print('Failed: %s' % (failed))
