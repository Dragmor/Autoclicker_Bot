import hashlib
from time import ctime
import ntplib
from datetime import datetime
from tkinter import messagebox

while True:
    ntp_client = ntplib.NTPClient()
    try:
        response = ntp_client.request('pool.ntp.org')
    except:
        try:
            response = ntp_client.request('uk.pool.ntp.org')
        except:
            try:
                response = ntp_client.request('0.de.pool.ntp.org')
            except:
                messagebox.showerror(title='Ошибка', message='Не удалось соединиться с сервером синхронизации времени!')
    # response = ntp_client.request('pool.ntp.org')

    date = datetime.strptime(ctime(response.tx_time), "%a %b %d %H:%M:%S %Y")

    formatted_date = date.strftime("%d%m%Y%S")

    md5 = hashlib.md5((formatted_date).encode()).hexdigest()

    out_string = ''
    for i in range(0, len(md5)):
        if i%3 != 0:
            out_string+=md5[i]
        else:
            if i//3 > len(formatted_date)-1:
                out_string+=md5[i]
            else:
                out_string+=formatted_date[i//3]
    print(out_string, end='\n\n\n\n\n')
    input()
