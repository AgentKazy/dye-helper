from time import sleep
# https://pynput.readthedocs.io/en/latest/keyboard.html#controlling-the-keyboard
from pynput import keyboard
from pynput.keyboard import Key, Controller
from ctypes import windll, create_unicode_buffer
import pyperclip
import os


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


def days():
    k = Controller()
    k.release('q')
    k.release(Key.alt)
    window_title = activeWindowTitle()
    while 'Detalhe do' in window_title:
        k.tap(Key.tab)
        k.tap(Key.tab)
        k.tap(Key.tab)
        k.tap(Key.tab)
        with k.pressed(Key.shift):
            k.tap(Key.end)
        k.tap(Key.menu)
        sleep(0.1)
        k.tap('c')
        sleep(0.2)
        read_numbers = pyperclip.paste()
        if read_numbers in {'3,00', '6,00', '7,00'}:
            k.type(str(8))
            pyperclip.copy('NONE')
            k.tap(Key.enter)
            return'''
        elif read_numbers == '3,00':
            k.type(str(8))
            pyperclip.copy('NONE')
            k.tap(Key.enter)
            return'''
        else:
            k.tap(Key.esc)
            return


def quantity():
    k = Controller()
    k.release('w')
    k.release(Key.alt)
    window_title = activeWindowTitle()
    while 'Detalhe do' in window_title:
        k.tap(Key.tab)
        k.tap(Key.tab)
        k.tap(Key.tab)
        k.tap(Key.tab)
        k.tap(Key.tab)
        k.tap(Key.tab)
        k.tap(Key.tab)
        k.tap(Key.tab)
        k.tap(Key.tab)
        k.tap(Key.tab)
        k.tap(Key.tab)
        k.tap(Key.tab)
        k.tap(Key.tab)
        k.tap(Key.tab)
        k.tap(Key.tab)
        k.tap(Key.tab)
        k.tap(Key.tab)
        with k.pressed(Key.shift):
            k.tap(Key.end)
        k.tap(Key.menu)
        sleep(0.1)
        k.tap('c')
        sleep(0.2)
        read_numbers = pyperclip.paste()
        if read_numbers in {'1000', '500'}:
            k.type(str(750))
            pyperclip.copy('NONE')
            k.tap(Key.enter)
            return
        else:
            k.tap(Key.esc)
            return


def copy_dye_values():
    k = Controller()
    if not window_title_exists():
        pass
    else:
        dye_values = []
        pyperclip.copy('dosorama')
        while True:
            while True:
                k.tap(Key.menu)
                k.tap('c')
                sleep(0.2)
                read_value = pyperclip.paste()
                # print('Test value 0: ' + read_value)
                if 'dosorama' not in read_value:
                    # print('Test value 1: ' + read_value)
                    break
            if 'NONE' in read_value:
                # print('Test value 2: ' + read_value)
                break
            else:
                values_str = str(read_value.split('\\')[0]).replace(",", ".")
                dye_values.append(float(values_str))
                pyperclip.copy('NONE')
                if test_mode is False:
                    # print('Test value 3: ' + read_value)
                    k.tap(Key.enter)
                    k.tap(Key.enter)
                else:
                    print('Dye value: ' + values_str, end='\r')
                    k.tap(Key.down)
        # Sum all quantities from list: https://stackoverflow.com/a/11344839
        dye_sum = float(sum(i for i in dye_values))
        # Check for Indifix PA
        k.tap(Key.esc)
        k.tap(Key.home)
        k.tap(Key.up)
        if test_mode is False:
            k.tap(Key.enter)
        k.tap(Key.menu)
        k.tap('c')
        sleep(0.2)
        indifix = pyperclip.paste()
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
        if test_mode is False:
            k.tap(Key.enter)
        if fabric in {'cotton', 'viscose'}:
            pyperclip.copy('dosorama')
            biotouch = pyperclip.paste()
            while True:
                k.tap(Key.menu)
                k.tap('c')
                sleep(0.2)
                biotouch = pyperclip.paste()
                if 'dosorama' not in biotouch:
                    break
            if biotouch not in {'2211', '2212'}:
                k.tap(Key.esc)
                k.tap(Key.right)
                k.tap(Key.right)
                k.tap(Key.right)
                if test_mode is False:
                    k.tap(Key.enter)
            else:
                reactive = True
                if fabric == 'cotton':
                    k.type('2211')  # Bio-touch CN
                elif fabric == 'viscose':
                    k.type('2212')  # Bio-touch NLG
                k.tap(Key.enter)
                if test_mode is False:
                    k.tap(Key.enter)
                else:
                    k.tap(Key.right)
                    k.tap(Key.right)
                    k.tap(Key.right)
        # Read and sum dye quantities
        total_dye_quantity = copy_dye_values()
        if fabric in {'cotton', 'viscose'}:
            salt_total, soda_ash_total = calculate_salt_soda(
                total_dye_quantity, fabric)
        elif fabric == 'acid':
            indigive_total = calculate_acid_donor(total_dye_quantity)
        # Reset position
        k.tap(Key.esc)
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
            k.type(str(indigive_total))
            k.tap(Key.enter)
        elif fabric in {'cotton', 'viscose'}:
            k.type(str(salt_total))
            k.tap(Key.enter)
            if half:
                k.type(str(int(soda_ash_total) / 2))
            else:
                k.type(str(soda_ash_total))
            k.tap(Key.enter)
        if reactive is True:
            k.tap(Key.esc)
            k.tap(Key.down)
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
    sleep(0.5)
    previous = pyperclip.paste()
    sleep(0.25)
    navigate_dosorama(fabric, half)
    if test_mode is True:
        print()
    pyperclip.copy(previous)
    return


def calculate_co():
    calculate_dyes('cotton', Key.f5, False)
    return


def calculate_cv():
    calculate_dyes('viscose', Key.f6, False)
    return


def calculate_half_co():
    calculate_dyes('cotton', Key.f7, True)
    return


def calculate_half_cv():
    calculate_dyes('viscose', Key.f8, True)
    return


def calculate_ac():
    calculate_dyes('acid', Key.f9, False)
    return


with keyboard.GlobalHotKeys({
     '<alt>+q': days,
     '<alt>+w': quantity,
     '<f5>': calculate_co,
     '<f6>': calculate_cv,
     '<f9>': calculate_ac}) as listener:
    listener.join()
