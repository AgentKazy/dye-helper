from ctypes import windll, create_unicode_buffer
import os
# https://pynput.readthedocs.io/en/latest/keyboard.html#controlling-the-keyboard
from pynput import keyboard
from pynput.keyboard import Key, Controller
import pyperclip
from time import sleep


os.system('cls')

test_mode = False

# ----------------- Do not edit below this line ----------------- #

if test_mode is True:
    window_config = set(['Excel', 'Go to'])
else:
    window_config = set(['Dosorama'])


# https://stackoverflow.com/a/58355052
def activeWindowTitle():
    hWnd = windll.user32.GetForegroundWindow()
    length = windll.user32.GetWindowTextLengthW(hWnd)
    buffer = create_unicode_buffer(length + 1)
    windll.user32.GetWindowTextW(hWnd, buffer, length + 1)
    return buffer.value if buffer.value else None


# https://stackoverflow.com/a/27620671/11192443
def window_title_exists():
    window_title = activeWindowTitle()
    if window_title:
        title_words = set(word for word in window_title.split())
        return window_config.intersection(title_words)


def calculate_salt_soda(dye_sum, fabric):
    if fabric in {'cotton'}:
        # https://thispointer.com/python-dictionary-with-multiple-values-per-key/
        salt_soda = {
            0.1: ['20', '6'],
            0.3: ['30', '8'],
            0.6: ['40', '10'],
            1.0: ['50', '15'],
            1.5: ['60', '20'],
            3.0: ['70', '20'],
            8.5: ['80', '20']}
    elif fabric in {'viscose'}:
        salt_soda = {
            0.1: ['15', '6'],
            0.3: ['20', '8'],
            0.6: ['30', '10'],
            1.0: ['40', '12'],
            1.5: ['50', '15'],
            3.0: ['60', '15'],
            5.0: ['70', '20'],
            9.0: ['80', '20']}
    max_values = ['90', '20']
    for quantity in salt_soda.keys():
        if dye_sum < quantity:
            if test_mode is True:
                print('Total salt calculated: '
                      + str(salt_soda[quantity][0]))
                print('Total soda ash calculated: '
                      + str(salt_soda[quantity][1]))
            return salt_soda[quantity][0], salt_soda[quantity][1]
    # If value is over 'salt_soda'
    if test_mode is True:
        print('Total salt calculated: '
              + str(max_values[0]))
        print('Total soda ash calculated: '
              + str(max_values[1]))
    return max_values[0], max_values[1]


def calculate_acid_donor(dye_sum):
    indigive = {
            0.3: '0.3',
            0.5: '0.5',
            0.75: '0.75',
            1.0: '1.0',
            1.5: '1.5'}
    max_value = '1.5'
    for quantity in indigive.keys():
        if dye_sum < quantity:
            if test_mode is True:
                print('Total acid donor calculated: '
                      + str(indigive[quantity]))
            return indigive[quantity]
    # If value is over 'indigive'
    if test_mode is True:
        print('Total acid donor calculated: ' + str(max_value))
    return max_value


def copy_dye_values():
    k = Controller()
    if not window_title_exists():
        pass
    else:
        dye_values = []
        while True:
            read_value = 'break_loop'
            while read_value == 'break_loop':
                with k.pressed(Key.ctrl):
                    k.tap('c')
                sleep(0.05)
                read_value = pyperclip.paste()
            # sleep(0.2)
            if 'NONE' in read_value:
                break
            dye_read = str(read_value.replace(' ', '')).replace(',', '.')
            dye_values.append(float(dye_read))
            pyperclip.copy('NONE')
            if test_mode is True:
                print('Dye value: ' + dye_read, end='\r')
            k.tap(Key.down)
        # Sum all quantities from list: https://stackoverflow.com/a/11344839
        dye_sum = float(sum(i for i in dye_values))
        # Check for Indifix PA
        k.tap(Key.home)
        k.tap(Key.up)
        indifix = 'break_loop'
        while indifix == 'break_loop':
            with k.pressed(Key.ctrl):
                k.tap('c')
            sleep(0.05)
            indifix = pyperclip.paste()
        # sleep(0.2)
        if '2026' not in indifix:
            pass
        else:
            indifix_value = float(dye_values[-1])
            if test_mode is True:
                print('Formula contains Indifix PA: '
                      + str((dye_values[-1])) + ' g/L')
            dye_sum = dye_sum - indifix_value
        if test_mode is True:
            print('Total dye quantity: ' + str(round(dye_sum, 3)))
        return dye_sum


def navigate_dosorama(fabric, half):
    k = Controller()
    reactive = False
    if not window_title_exists():
        if test_mode is True:
            print('No required windows currently active.')
        pass
    else:
        # Reset position
        k.tap(Key.esc)
        with k.pressed(Key.ctrl):
            k.tap(Key.home)
        # Move to first dye
        if fabric == 'acid':
            k.tap(Key.right)
            k.tap(Key.right)
            k.tap(Key.right)
        k.tap(Key.down)
        k.tap(Key.down)
        k.tap(Key.down)
        # Check for Bio-touch
        if fabric in {'cotton', 'viscose'}:
            k.tap(Key.down)
        if fabric in {'cotton', 'viscose'}:
            biotouch = 'break_loop'
            pyperclip.copy('break_loop')
            while biotouch == 'break_loop':
                with k.pressed(Key.ctrl):
                    k.tap('c')
                # print('before = ' + str(biotouch))
                biotouch = pyperclip.paste()
                sleep(0.005)
                # print('after = ' + str(biotouch))
            # sleep(0.2)
            if biotouch not in {'2211', '2212'}:
                k.tap(Key.esc)
                k.tap(Key.right)
                k.tap(Key.right)
                k.tap(Key.right)
            else:
                reactive = True
                if fabric == 'cotton':
                    k.type('2211')  # Bio-touch CN
                elif fabric == 'viscose':
                    k.type('2212')  # Bio-touch NLG
                pyperclip.copy('break_loop')
                k.tap(Key.enter)
                sleep(0.05)
                k.tap(Key.enter)
        # Read and sum dye quantities
        total_dye_quantity = copy_dye_values()
        if fabric in {'cotton', 'viscose'}:
            salt_total, soda_ash_total = calculate_salt_soda(
                total_dye_quantity, fabric)
        elif fabric == 'acid':
            indigive_total = calculate_acid_donor(total_dye_quantity)
        # Reset position
        with k.pressed(Key.ctrl):
            k.tap(Key.home)
        # Move to Salt/Acid donor values
        k.tap(Key.right)
        k.tap(Key.right)
        k.tap(Key.right)
        k.tap(Key.down)
        k.tap(Key.down)
        # Enter salt and soda ash/acid donor totals
        if fabric == 'acid':
            k.type('0')
            with k.pressed(Key.ctrl):
                k.type('a')
            k.type(str(indigive_total.replace(' ', '')))
            k.tap(Key.enter)
        elif fabric in {'cotton', 'viscose'}:
            k.type('0')
            with k.pressed(Key.ctrl):
                k.type('a')
            k.type(str(salt_total.replace(' ', '')))
            k.tap(Key.enter)
            sleep(0.05)
            k.type('0')
            with k.pressed(Key.ctrl):
                k.type('a')
            '''sleep(0.2)
            with k.pressed(Key.ctrl):
                k.type('a')'''
            if half:
                k.type(str(int(soda_ash_total.replace(' ', '')) / 2))
                # sleep(0.2)
            else:
                k.type(str(soda_ash_total.replace(' ', '')))
                # sleep(0.2)
            k.tap(Key.enter)
            sleep(0.05)
        if reactive is True:
            k.tap(Key.enter)
        return


def calculate_dyes(fabric, key, half):
    k = Controller()
    k.release(key)
    k.release(Key.ctrl)
    k.release(Key.shift)
    k.release(Key.alt)
    if test_mode is True:
        key_split = str(key).split('.')[1]
        print(key_split.upper() + ' --> Calculating ' + fabric + ' values')
    sleep(0.25)
    previous = pyperclip.paste()
    sleep(0.25)
    navigate_dosorama(fabric, half)
    if test_mode is True:
        print()
    pyperclip.copy(previous)
    return


def calculate_co():
    calculate_dyes('cotton', Key.f1, False)
    return


def calculate_cv():
    calculate_dyes('viscose', Key.f2, False)
    return


def calculate_half_co():
    calculate_dyes('cotton', Key.f3, True)
    return


def calculate_half_cv():
    calculate_dyes('viscose', Key.f4, True)
    return


def calculate_ac():
    calculate_dyes('acid', Key.f5, False)
    return


with keyboard.GlobalHotKeys({
     '<f1>': calculate_co,
     '<f2>': calculate_cv,
     '<f3>': calculate_half_co,
     '<f4>': calculate_half_cv,
     '<f5>': calculate_ac}) as listener:
    listener.join()
