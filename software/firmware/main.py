"""
MicroPython baseline for Seeed XIAO RP2040 Volume Mixer
Pinout and hardware mapping per schematic and details.txt
Features:
- Reads 5 analog faders (RV1-RV4 via RP2040 ADC, RV5 via ADS1115)
- Sets up MIDI CC output for each fader
- Handles 5 SSD1306 OLED displays (each on its own TCA9548A I2C channel)
- Handles push buttons
- Controls SK6812 LEDs
- Receives data from desktop app via USB serial
"""

import machine
import utime
import ustruct
from machine import Pin, ADC, I2C


# --- Pin Definitions (from schematic) ---
# I2C Bus
I2C_SDA = 6  # GPIO6
I2C_SCL = 7  # GPIO7

# TCA9548A I2C Multiplexer
TCA9548A_ADDR = 0x70  # Base address, adjust A0/A1/A2 as needed
SSD1306_CHANNELS = [0, 1, 2, 3, 4]  # CH0-CH4 for 5 displays

# ADS1115 (direct on I2C bus)
ADS1115_ADDR = 0x48

# Faders (Potentiometers)
FADER_ADC_PINS = [26, 27, 28, 29]  # RV1-RV4: GPIO26-29 (ADC0-ADC3)
FADER_ADS_CHANNEL = 0  # RV5: ADS1115 channel (A0)

# Push Buttons
BUTTON_PINS = [0, 1]  # SW1, SW2

# SK6812 LEDs
LED_PIN = 4
LED_COUNT = 8


# --- MIDI CC Setup ---
MIDI_CHANNEL = 0
FADER_CC = [20, 21, 22, 23, 24]  # Example CC numbers for faders


# --- I2C Setup ---
i2c = I2C(0, scl=Pin(I2C_SCL), sda=Pin(I2C_SDA), freq=400000)


# --- ADC Setup ---
faders = [ADC(Pin(pin)) for pin in FADER_ADC_PINS]  # RV1-RV4


# --- ADS1115 Setup ---
class ADS1115:
    def __init__(self, i2c, addr=ADS1115_ADDR):
        self.i2c = i2c
        self.addr = addr
    def read(self, channel=0):
        # Minimal single-ended read (for RV5)
        config = 0x4000 | (channel << 12) | 0x8583
        self.i2c.writeto_mem(self.addr, 0x01, ustruct.pack('>H', config))
        utime.sleep_ms(2)
        raw = self.i2c.readfrom_mem(self.addr, 0x00, 2)
        val = ustruct.unpack('>h', raw)[0]
        return max(0, min(65535, val + 32768))
ads = ADS1115(i2c, ADS1115_ADDR)


# --- Button Setup ---
buttons = [Pin(pin, Pin.IN, Pin.PULL_UP) for pin in BUTTON_PINS]


# --- LED Setup (SK6812) ---
# Placeholder: use neopixel library if available
try:
    import neopixel
    leds = neopixel.NeoPixel(Pin(LED_PIN), LED_COUNT)
except ImportError:
    leds = None


# --- MIDI CC Send ---
def send_midi_cc(cc, value, channel=0):
    # Send MIDI CC over USB serial
    msg = bytes([0xB0 | channel, cc, value])
    machine.UART(0, 31250).write(msg)  # If using UART MIDI
    # For USB MIDI, use pyb.USB_VCP or custom implementation


# --- USB Serial Receive ---
# Use machine.UART(1, ...) or pyb.USB_VCP for desktop comm
try:
    usb = machine.UART(1, 115200)
except:
    usb = None

def read_usb_data():
    if usb and usb.any():
        return usb.read()
    return None


# --- TCA9548A Multiplexer Helper ---
def tca9548a_select_channel(i2c, addr, channel):
    # Selects the given channel (0-7) on the TCA9548A
    i2c.writeto(addr, bytes([1 << channel]))

# --- Main Loop ---
def set_led_bar(val, count=LED_COUNT):
    # Set LED bar: fill LEDs proportional to val (0-127)
    if leds:
        num_on = int((val / 127.0) * count)
        for i in range(count):
            if i < num_on:
                leds[i] = (8, 217, 214)  # Cyan
            else:
                leds[i] = (32, 32, 32)   # Dim/off
        leds.write()

def main():
    last_fader = [0] * 5
    while True:
        # Read faders (RV1-RV4: RP2040 ADC, RV5: ADS1115)
        fader_vals = [fader.read_u16() >> 9 for fader in faders]  # 0-127
        fader_vals.append(ads.read(FADER_ADS_CHANNEL) >> 9)

        # Send MIDI CC if changed
        for i, val in enumerate(fader_vals):
            if val != last_fader[i]:
                send_midi_cc(FADER_CC[i], val)
                last_fader[i] = val

        # LED bar: show main volume (first fader)
        set_led_bar(fader_vals[0], LED_COUNT)

        # Read buttons
        button_states = [not btn.value() for btn in buttons]

        # Example: Update each OLED display (SSD1306) via TCA9548A
        for ch in SSD1306_CHANNELS:
            tca9548a_select_channel(i2c, TCA9548A_ADDR, ch)
            # TODO: Add SSD1306 update code here (e.g., display fader value)

        # Handle USB data from desktop app
        data = read_usb_data()
        if data:
            # Parse and update faders/screens/leds as needed
            pass  # Implement protocol as needed

        utime.sleep_ms(10)

if __name__ == "__main__":
    main()
