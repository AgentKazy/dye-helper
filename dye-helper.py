from ctypes import windll, create_unicode_buffer
import os
import keyboard as kb
import pyperclip as ppc
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
    return max_values[0].replace(' ', ''), max_values[1].replace(' ', '')


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
            return indigive[quantity].replace(' ', '')
    # If value is over 'indigive'
    if test_mode is True:
        print('Total acid donor calculated: ' + str(max_value))
    return max_value.replace(' ', '')


def copy_dye_values():
    if not window_title_exists():
        pass
    else:
        dye_values = []
        while True:
            read_value = 'break_loop'
            while read_value == 'break_loop':
                if test_mode is True:
                    print("Before copying: " + read_value)
                kb.send('ctrl+c')
                sleep(0.05)
                read_value = str(
                    ppc.paste().replace(' ', '')).replace(',', '.')
                if test_mode is True:
                    print("After copying: " + read_value)
            # sleep(0.2)
            if '.' not in read_value:
                break
            if test_mode is True:
                print('Read value :' + read_value)
            dye_values.append(float(read_value))
            ppc.copy('NONE')
            if test_mode is True:
                print('Read value: ' + read_value, end='\r')
                print('Dye value: ' + str(
                    float(sum(i for i in dye_values))), end='\r')
            kb.send('down')
        # Sum all quantities from list: https://stackoverflow.com/a/11344839
        dye_sum = float(sum(i for i in dye_values))
        # Check for Indifix PA
        kb.send('home, up')
        indifix = 'break_loop'
        while indifix == 'break_loop':
            kb.send('ctrl+c')
            sleep(0.05)
            indifix = ppc.paste()
        # sleep(0.2)
        if '2026' not in indifix:
            pass
        else:
            indifix_value = float(dye_values[-1])
            if test_mode is True:
                print(dye_values[-1])
            if test_mode is True:
                print('Formula contains Indifix PA: '
                      + str((dye_values[-1])) + ' g/L')
                print('Dye sum before removing Indifix: ' + str(dye_sum))
            dye_sum = dye_sum - indifix_value
        if test_mode is True:
            print('Indifix value: ' + str(indifix_value))
            print('Total dye quantity (after Indifix): ' + (
                str(round(dye_sum, 3))))
        return dye_sum


def navigate_dosorama(fabric, half):
    reactive = False
    if not window_title_exists():
        if test_mode is True:
            print('No required windows currently active.')
        pass
    else:
        # Reset position
        kb.send('esc, ctrl+home')
        # Move to first dye
        if fabric == 'acid':
            kb.send('right, right, right')
        kb.send('down, down, down')
        # Check for Bio-touch
        if fabric in {'cotton', 'viscose'}:
            kb.send('down')
        if fabric in {'cotton', 'viscose'}:
            biotouch = 'break_loop'
            ppc.copy('break_loop')
            while biotouch == 'break_loop':
                if test_mode is True:
                    print('before = ' + str(biotouch))
                kb.send('ctrl+c')
                sleep(0.05)
                biotouch = ppc.paste()
                if test_mode is True:
                    print('after = ' + str(biotouch))
            # sleep(0.2)
            if test_mode is True:
                print('Biotouch value is: ' + str(biotouch))
            if not biotouch.startswith('221'):
                if test_mode is True:
                    print('No bio-touch.')
                kb.send('esc, right, right, right')
            else:
                reactive = True
                if fabric == 'cotton':
                    kb.write('2211')  # Bio-touch CN
                elif fabric == 'viscose':
                    kb.write('2212')  # Bio-touch NLG
                ppc.copy('break_loop')
                kb.send('enter')
                if test_mode is True:
                    kb.send('up, right, right, right, down')
                sleep(0.05)
                kb.send('enter')
                sleep(0.05)
        # Read and sum dye quantities
        total_dye_quantity = copy_dye_values()
        if fabric in {'cotton', 'viscose'}:
            salt_total, soda_ash_total = calculate_salt_soda(
                total_dye_quantity, fabric)
        elif fabric == 'acid':
            indigive_total = calculate_acid_donor(total_dye_quantity)
        # Reset position & move to Salt/Acid donor values
        kb.send('ctrl+home, right, right, right, down, down')
        # Enter salt and soda ash/acid donor totals
        if fabric == 'acid':
            # kb.write('0')
            # kb.send('ctrl+a')
            kb.write(str(indigive_total))
            kb.send('enter')
        elif fabric in {'cotton', 'viscose'}:
            # kb.write('0')
            # kb.send('ctrl+a')
            kb.write(str(salt_total))
            kb.send('enter')
            sleep(0.05)
            # kb.write('0')
            # kb.send('ctrl+a')
            # sleep(0.2)
            if half:
                kb.write(str(int(soda_ash_total) / 2))
                # sleep(0.2)
            else:
                kb.write(str(soda_ash_total))
                # sleep(0.2)
            kb.send('enter')
            sleep(0.05)
        if reactive is True:
            kb.send('enter')
        return


def calculate_dyes(fabric, key, half):
    if test_mode is True:
        print(key.upper() + ' --> Calculating ' + fabric + ' values')
    # sleep(0.25)
    previous = ppc.paste()
    sleep(0.25)
    navigate_dosorama(fabric, half)
    if test_mode is True:
        print()
    ppc.copy(previous)
    return


kb.add_hotkey('f1', calculate_dyes, args=('cotton', 'f1', False))
kb.add_hotkey('f2', calculate_dyes, args=('viscose', 'f2', False))
kb.add_hotkey('f3', calculate_dyes, args=('cotton', 'f3', True))
kb.add_hotkey('f4', calculate_dyes, args=('viscose', 'f4', True))
kb.add_hotkey('f5', calculate_dyes, args=('acid', 'f5', False))


kb.wait()
