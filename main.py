import tkinter as tk
from tkinter import *
from tkinter import filedialog
from tkinter import scrolledtext

from PIL import Image, ImageFont, ImageDraw
import pandas as pd
import textwrap
import smtplib
import os
import re
from io import BytesIO
import threading
import math

from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart


FILE_PATH = ''
SAVE_PATH = ''
BAL_GRAPHIC = 'BM Graphic.jpg'
KISHORE_GRAPHIC = 'KM Graphic.jpg'
GRAPHIC_PATH = BAL_GRAPHIC
EMAIL_SUBJECT = 'BK Vision 2022'
SMTP_SERVER = None
SENDING_EMAILS_BUTTON = None
graphic_type = None

def get_file_path(file_path_field):
    file_path = filedialog.askopenfilename(initialdir=os.getcwd())
    file_path_field.delete(0, 'end')
    file_path_field.insert(0, file_path)


def get_image_save_path(image_save_field):
    save_path = filedialog.askdirectory(initialdir=os.getcwd())
    image_save_field.delete(0, 'end')
    image_save_field.insert(0, save_path)


def get_graphic_path(graphic_path_field):
    image_path = filedialog.askdirectory(initialdir=os.getcwd())
    graphic_path_field.delete(0, 'end')
    graphic_path_field.insert(0, image_path)


def on_change_file_path_field(sv):
    global FILE_PATH
    FILE_PATH = sv.get()


def on_change_image_save_path_field(sv):
    global SAVE_PATH
    SAVE_PATH = sv.get()


def file_select_UI(root, canvas):
    label = tk.Label(root,
                     text='Select the file with the answers',
                     font=('helvetica', 10),
                    )

    canvas.create_window(105, 60, window=label)

    sv = StringVar()
    sv.trace("w",
             lambda name,
             index, mode,
             sv=sv:
             on_change_file_path_field(sv)
             )

    file_path_field = tk.Entry(root, textvariable=sv, width=40)
    canvas.create_window(137, 80, window=file_path_field)
    canvas.pack()

    button = tk.Button(
        text='Browse',
        command = lambda: get_file_path(file_path_field),
    )

    canvas.create_window(283, 80, window=button)


def image_save_UI(root, canvas):
    label = tk.Label(root,
                     text='Folder where you want to save the graphics',
                     font=('helvetica', 10),
                    )

    canvas.create_window(140, 120, window=label)

    sv = StringVar()
    sv.trace("w",
             lambda name,
             index, mode,
             sv=sv:
             on_change_image_save_path_field(sv)
             )

    image_save_field = tk.Entry(root, textvariable=sv, width=40)
    canvas.create_window(137, 140, window=image_save_field)
    canvas.pack()

    button = tk.Button(
        text='Browse',
        command = lambda: get_image_save_path(image_save_field),
    )

    canvas.create_window(283, 140, window=button)


def select_graphic(value):
    global GRAPHIC_PATH
    if(value == 0):
        GRAPHIC_PATH = BAL_GRAPHIC
    else:
        GRAPHIC_PATH = KISHORE_GRAPHIC


def graphic_select_UI(root, canvas):
    radioGroup = tk.LabelFrame(root, text = "Select graphic type")
    radioGroup.pack()

    # Radio variable
    global graphic_type
    graphic_type = IntVar()
    graphic_type.set(0)

    # Create two radio buttons
    bal = Radiobutton(radioGroup, text = "Bal", variable = graphic_type,
                      value = 0, command= lambda: select_graphic(0))
    bal.pack(anchor=W)

    kishore = Radiobutton(radioGroup, text = "Kishore", variable = graphic_type,
                          value = 1, command= lambda: select_graphic(1))
    kishore.pack(anchor=W)

    canvas.create_window(73, 200, window=radioGroup)


def create_log(root, canvas):
    log_box = scrolledtext.ScrolledText(
        root,
        wrap = tk.WORD,
        state=DISABLED,
        height=15,
        width=75,
    )

    canvas.create_window(630, 180, window=log_box)

    return log_box


def email_submission(root, canvas, log):
    global SENDING_EMAILS_BUTTON

    SENDING_EMAILS_BUTTON = tk.Button(
        text='Send email',
        state = tk.NORMAL,
        command = lambda: spawn_submission_thread(root, log)
    )

    canvas.create_window(50, 260, window=SENDING_EMAILS_BUTTON)

def check_email(email, log, emailIndex, imgObj):
    try:
        regex = '^[a-z0-9]+[\._]?[a-z0-9]+[@]\w+[.]\w{2,3}$'
        if(re.search(regex,email)):
            send_mail(imgObj, email, log)
            log['state'] = tk.NORMAL
            log.insert(END, f'Email {emailIndex} sent to: {email}\n')
            log['state'] = tk.DISABLED
        else:
            log['state'] = tk.NORMAL
            log.insert(END, f'Tried Email {emailIndex}: {email}.  Failed Incorrect Email\n')
            log['state'] = tk.DISABLED
    except TypeError:
        emailIndex += 1
        log['state'] = tk.NORMAL
        log.insert(END, f'Failed Incorrect Email: Check Excel Row {emailIndex}\n')
        log['state'] = tk.DISABLED


def parse_answers(root, log):
    if(not FILE_PATH):
        log['state'] = tk.NORMAL
        log.insert(END, "Please ensure you have selected file path\n")
        log['state'] = tk.DISABLED
        return

    if(not SAVE_PATH):
        log['state'] = tk.NORMAL
        log.insert(END, "Please ensure you have a save path\n")
        log['state'] = tk.DISABLED
        return

    global SENDING_EMAILS_BUTTON_STATE
    SENDING_EMAILS_BUTTON_STATE = DISABLED
    root.update()

    try:
        data = pd.read_excel(FILE_PATH)
    except:
        log['state'] = tk.NORMAL
        log.insert(END, f"Cannot open the answers file at path: {FILE_PATH}\n")
        log['state'] = tk.DISABLED
        return

    rows = data.to_numpy()

    try:
        font_120 = ImageFont.truetype(f"{os.getcwd()}/fonts/rockwen.ttf", 120)
        font_140 = ImageFont.truetype(f"{os.getcwd()}/fonts/rockwen.ttf", 140)
        font_208 = ImageFont.truetype(f"{os.getcwd()}/fonts/rockb.ttf", 208)
    except:
        log['state'] = tk.NORMAL
        log.insert(END, f"Cannot open all the fonts\n")
        log['state'] = tk.DISABLED
        return

    log['state'] = tk.NORMAL
    log.insert(END, 'Sending emails...\n')
    log.insert(END, '\n')
    log['state'] = tk.DISABLED

    emailIndex = 1
    for row in rows:
        image = Image.open(f'{GRAPHIC_PATH}')

        image_width, image_height = image.size
        drawable_image = ImageDraw.Draw(image)

        sendEmail = True

        for answer_index in range(1, len(row)-1, 1):
            try:
                answer = row[answer_index]

                if(answer_index == 2):
                    answer = answer.upper()

                if(answer_index == 4):
                    quarter_prints(row, drawable_image, answer_index, font_120)
                elif(answer_index > 14):
                    if(answer_index == 15):
                        y_pos = 9313 - 50
                    else:
                        y_pos = 10073 - 50

                    culture_change_prints(answer,
                                          drawable_image,
                                          image_width,
                                          y_pos,
                                          font_120)
                else:
                    fonts = (font_120, font_140, font_208)

                    colour, x_pos, y_pos, font = text_print_options(answer,
                                                                    answer_index,
                                                                    image_width,
                                                                    image_height,
                                                                    fonts,
                                                                    drawable_image)

                    drawable_image.text((x_pos, y_pos),
                                        str(answer),
                                        font=font,
                                        fill=colour)

            except (ValueError, TypeError, AttributeError):
                emailIndex += 1
                columnIdx = answer_index + 1
                log.insert(END, f'Failed Incorrect Column: Check Excel Row {emailIndex}, Column {answer_index}\n')
                sendEmail = False

        if sendEmail:
            try:
                image_save = f"{SAVE_PATH}/bkvision_{emailIndex}_{row[1]}.jpg"
                image.save(image_save, format="JPEG")
            except:
                log.insert(END, f"Could not save image for {row[17]}/n")

            stream = BytesIO()
            image.save(stream, format="png")
            stream.seek(0)
            imgObj = stream.read()

            email_to = row[17]
            check_email(email_to, log, emailIndex, imgObj)

        emailIndex += 1
        root.update()

    log['state'] = tk.NORMAL
    log.insert(END, '\n')
    log.insert(END, f'All Emails Are Sent\n')
    log['state'] = tk.DISABLED

    root.update()

    SENDING_EMAILS_BUTTON_STATE = NORMAL
    root.update()


def quarter_prints(row, drawable_image, answer_index, font_120):
    increment = get_increment(row[answer_index], row[5])
    quarter = row[answer_index]

    x_pos = 1654
    y_pos = 3332
    y_increase = 366

    for quarter_index in range(0, 4, 1):
        quarter += increment

        drawable_image.text((x_pos, y_pos),
                            str(quarter),
                            font=font_120,
                            fill='rgb(0, 0, 0)')

        y_pos += y_increase

    y_pos = 3332
    x_pos = 4386

    for quarter_index in range(0, 4, 1):
        quarter += increment

        if(quarter_index == 3):
            quarter = row[5]

        drawable_image.text((x_pos, y_pos),
                            str(quarter),
                            font=font_120,
                            fill='rgb(0, 0, 0)')

        y_pos += y_increase


def culture_change_prints(answer, drawable_image, image_width, y_pos, font_120):
    answer = str(answer)

    wrapper = textwrap.TextWrapper(width=80)
    line_list = wrapper.wrap(text=answer)

    for line in line_list:
        text_width, text_height = drawable_image.textsize(line, font=font_120)
        x_pos = get_center_position(image_width, text_width)

        drawable_image.text((x_pos, y_pos),
                            line,
                            font=font_120,
                            fill='rgb(0, 0, 0)')
        y_pos += 130


def text_print_options(answer, answer_index, image_width, image_height, fonts, drawable_image):
    colour_black = 'rgb(0, 0, 0)'
    colour_white = 'rgb(245, 214, 219)'

    font_120, font_140, font_208 = fonts
    answer = str(answer)

    if(answer_index == 1):
        font = font_140
        text_width, text_height = drawable_image.textsize(answer, font=font)

        x_pos = get_starting_position(text_width, 4771)
        return (colour_white, x_pos, 1428, font)

    elif(answer_index == 2):
        font = font_208
        text_width, text_height = drawable_image.textsize(answer, font=font)
        x_pos = get_center_position(image_width, text_width)
        return (colour_white, x_pos, 1115, font)

    elif(answer_index == 3):
        font = font_140
        return (colour_white, 697, 1428, font)

    elif(answer_index == 5):
        font = font_140
        text_width, text_height = drawable_image.textsize(answer, font=font)
        x_pos = get_center_position(image_width, text_width)
        return (colour_black, x_pos, 2699, font)

    elif(answer_index == 6):
        font = font_140
        text_width, text_height = drawable_image.textsize(answer, font=font)
        x_pos = get_center_position(image_width, text_width)
        return (colour_black, x_pos, 5568, font)

    elif(answer_index == 7):
        font = font_140
        text_width, text_height = drawable_image.textsize(answer, font=font)
        x_pos = get_section_center(1365, text_width)
        return (colour_black, x_pos, 6097, font)

    elif(answer_index == 8):
        font = font_140
        text_width, text_height = drawable_image.textsize(answer, font=font)
        x_pos = get_section_center(1365, text_width)
        return (colour_black, x_pos, 6787, font)

    elif(answer_index == 9):
        font = font_140
        text_width, text_height = drawable_image.textsize(answer, font=font)
        x_pos = get_section_center(1365, text_width)
        return (colour_black, x_pos, 7480, font)

    elif(answer_index == 10):
        font = font_140
        text_width, text_height = drawable_image.textsize(answer, font=font)
        x_pos = get_section_center(1365, text_width)
        return (colour_black, x_pos, 8152, font)

    elif(answer_index == 11):
        font = font_140
        text_width, text_height = drawable_image.textsize(answer, font=font)
        x_pos = get_section_center(4097, text_width)
        return (colour_black, x_pos, 6097, font)

    elif(answer_index == 12):
        font = font_140
        text_width, text_height = drawable_image.textsize(answer, font=font)
        x_pos = get_section_center(4097, text_width)
        return (colour_black, x_pos, 6787, font)

    elif(answer_index == 13):
        font = font_140
        text_width, text_height = drawable_image.textsize(answer, font=font)
        x_pos = get_section_center(4097, text_width)
        return (colour_black, x_pos, 7480, font)

    elif(answer_index == 14):
        font = font_140
        text_width, text_height = drawable_image.textsize(answer, font=font)
        x_pos = get_section_center(4097, text_width)
        return (colour_black, x_pos, 8152, font)


def get_increment(initial, final):
    return (int(final) - int(initial)) // 8


def get_starting_position(text_width, end_x):
    return end_x - text_width


def get_section_center(section_center_x, text_width):
    return section_center_x - (text_width/2)


def get_center_position(image_width, text_width):
    x_pos = (image_width/2 - text_width/2)//1
    return x_pos


def send_mail(imgObj, email_to, log):
    if(not SMTP_SERVER):
        return

    email_from = 'bkvision22@gmail.com'

    msg = MIMEMultipart()
    msg['Subject'] = EMAIL_SUBJECT
    msg['From'] = email_from
    msg['To'] = email_to

    msgImg=MIMEImage(imgObj)
    msg.attach(msgImg)

    SMTP_SERVER.sendmail(email_from, email_to, msg.as_string())


def setup_smtp():
    global SMTP_SERVER
    SMTP_SERVER = smtplib.SMTP('smtp.gmail.com', 587)
    SMTP_SERVER.ehlo()
    SMTP_SERVER.starttls()
    SMTP_SERVER.ehlo()
    SMTP_SERVER.login('bkvision22@gmail.com', 'Bkvision2022')


def check_thread(thread, root):
    global SENDING_EMAILS_BUTTON

    if not thread.is_alive():
        SENDING_EMAILS_BUTTON['state'] = tk.NORMAL
        SENDING_EMAILS_BUTTON['relief'] = tk.RAISED
    else:
        SENDING_EMAILS_BUTTON['state'] = tk.DISABLED
        SENDING_EMAILS_BUTTON['relief'] = tk.SUNKEN

        # check again after 300ms
        root.after(300, lambda: check_thread(thread, root))


def spawn_submission_thread(root, log):
    SENDING_EMAILS_BUTTON_STATE = DISABLED
    new_thread = threading.Thread(target=parse_answers, args=(root, log))
    new_thread.start()

    check_thread(new_thread, root)


def create_gui():
    root= tk.Tk()
    setup_smtp()

    root.resizable(width = False, height = False)
    icon = tk.PhotoImage(file = "logo.png")
    root.iconphoto(False, icon)
    root.wm_title('BK Vision 2020')

    canvas_width = 980
    canvas_height = 350

    canvas = tk.Canvas(root, width = canvas_width, height = canvas_height)
    canvas.pack()

    log = create_log(root, canvas)
    file_select_UI(root, canvas)
    image_save_UI(root, canvas)
    graphic_select_UI(root, canvas)
    email_submission(root, canvas, log)

    root.mainloop()

print("Please wait while the program builds...")
create_gui()

# THINGS TO DO:
# - testing
# - check GUI on different screens
# - find new bundler (eg. not PyInstaller)
