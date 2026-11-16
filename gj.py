from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
import os

NAME = "YOUR FULL NAME"
MAT = "23CK033419"
COURSE = "Electronics Project Proposal"
DEPT = "Department of ________"

NANO = ("Arduino Nano (ATmega328P)", "5V microcontroller board", 1, 8000)
COMMON = [("Resistors", "Assorted 1/4W", 20, 15), ("Capacitors", "Assorted", 8, 60),
          ("LEDs", "5mm", 4, 30), ("BC547 transistors", "NPN driver", 3, 50),
          ("1N4007 diodes", "Flyback", 4, 30), ("Jumper wires and header pins", "", 1, 1500),
          ("Vero board", "Single sided", 1, 1500)]

PROJECTS = [
{
"slug": "GenGuard",
"title": "GenGuard: Generator Runtime, Temperature and Fuel Monitor with Automatic Overheat Shutdown",
"intro": [
 "Small petrol generators run for hours in homes, shops and hostels. Owners rarely know how long a generator has run since its last oil change, how much fuel is left, or whether it is overheating.",
 "GenGuard attaches to a generator, measures runtime, engine-area temperature and fuel level, warns when service is due, and cuts the engine kill circuit if temperature gets dangerous."],
"summary": [
 "Problem: missed servicing, surprise fuel-outs and overheating damage.",
 "Solution: Arduino-based monitor with LCD, buzzer and kill relay.",
 "Unique element: a service countdown stored in EEPROM so it survives power loss.",
 "Components reused from lower levels: relay and transistor driver, diodes, regulator, LEDs, buzzer."],
"principle": [
 "1. A DS18B20 temperature probe is fixed near the engine. The Nano reads it over one-wire every second.",
 "2. A float or ultrasonic sensor reads fuel level and the Nano converts it to a percentage.",
 "3. The Nano detects that the generator is running by sensing output voltage through a ZMPT101B module.",
 "4. While running, it adds elapsed time to a runtime counter and saves it to EEPROM every minute.",
 "5. At 50 running hours, the LCD shows 'SERVICE DUE' and the buzzer beeps at start-up.",
 "6. If temperature exceeds the set limit for 10 seconds, a BC547 energizes a relay that grounds the engine kill wire.",
 "7. A push button resets the service counter after maintenance."],
"sim": ["Wokwi or Proteus: DS18B20 as a temperature slider, potentiometer for fuel level, switch for 'running'.",
        "Speed time up by 60x to test the service counter and EEPROM write.",
        "Verify the relay triggers only after 10 s over limit."],
"nets": [
 ("DS18B20 data","Nano D2 + 4.7k pull-up to 5V","Temperature"),
 ("Fuel sensor / pot","Nano A0","Fuel level"),
 ("ZMPT101B OUT","Nano A1","Running detect"),
 ("I2C LCD SDA, SCL","Nano A4, A5","Display"),
 ("Nano D8","BC547 base via 1k","Kill driver"),
 ("BC547 collector","Relay coil + 1N4007","Kill relay"),
 ("Nano D9","Buzzer","Alarm"),
 ("Nano D3","Reset push button","Service reset"),
 ("12V adapter","LM7805 to Nano 5V pin","Supply")],
"beme": [NANO,
 ("DS18B20 waterproof probe","Temperature",1,2500),("Ultrasonic sensor or float","Fuel level",1,2500),
 ("ZMPT101B module","Mains sense, isolated",1,3500),("16x2 I2C LCD","Display",1,3500),
 ("5V relay module","10A",1,1500),("Buzzer","Active 5V",1,250),("LM7805","Regulator",1,200),
 ("Push buttons","Tactile",2,50),("Enclosure","Plastic box",1,3000)] + COMMON
},
{
"slug": "GasGauge",
"title": "GasGauge: LPG Cylinder Weight Monitor with Leak Detection and Cut-Off Alarm",
"intro": [
 "Households and food vendors cook with LPG and often discover the cylinder is empty halfway through cooking. Gas leaks are also a real fire risk.",
 "GasGauge sits under the cylinder, weighs it, calculates remaining gas and estimated days left, and sounds an alarm if a leak is detected."],
"summary": [
 "Problem: running out of gas unexpectedly and undetected leaks.",
 "Solution: load cell weighing platform and gas sensor on an Arduino Nano.",
 "Unique element: days-remaining estimate from daily usage history.",
 "Safety: a relay output can drive a solenoid valve or an extractor fan."],
"principle": [
 "1. A 50 kg load cell and HX711 amplifier measure total weight.",
 "2. The Nano subtracts the empty cylinder weight (tare) and converts the remainder to a gas percentage.",
 "3. Daily use is the weight drop over 24 hours. The Nano divides remaining gas by average daily use to show days left.",
 "4. An MQ-6 sensor reads LPG concentration through an analog pin.",
 "5. If the reading passes the limit, the buzzer sounds, a red LED lights and the relay turns on the extractor fan.",
 "6. The LCD shows weight, percent, days left and leak status.",
 "7. Low gas (below 15%) triggers a slow beep as a refill reminder."],
"sim": ["Proteus or Wokwi: replace HX711 with a potentiometer scaled to weight.",
        "Use a slider for the MQ-6 analog output.",
        "Check thresholds and the relay for leak and low-gas conditions."],
"nets": [
 ("Load cell E+, E-, A+, A-","HX711","Weight signal"),
 ("HX711 DT, SCK","Nano D4, D5","Digital read"),
 ("MQ-6 AOUT","Nano A0","Gas level"),
 ("I2C LCD","Nano A4, A5","Display"),
 ("Nano D7","BC547 to relay","Fan or valve"),
 ("Nano D8","Buzzer","Alarm"),
 ("Nano D9, D10","Green and red LEDs via 330R","Status"),
 ("5V adapter","Nano","Supply")],
"beme": [NANO,
 ("Load cell 50kg and HX711","Weighing",1,5000),("MQ-6 gas sensor module","LPG",1,3000),
 ("16x2 I2C LCD","Display",1,3500),("5V relay module","",1,1500),("Buzzer","Active 5V",1,250),
 ("Wood or acrylic platform","Weighing base",1,3000),("Enclosure","",1,2500),("5V adapter","",1,2000)] + COMMON
},
{
"slug": "ColdChain",
"title": "ColdChain Guard: Outage-Proof Temperature Logger for Insulin and Vaccine Storage",
"intro": [
 "Clinics and households store insulin and vaccines in refrigerators. During power cuts the temperature rises and the user may not know how long the medicine was exposed.",
 "ColdChain Guard measures fridge temperature on its own battery, logs time spent out of range, and warns the user."],
"summary": [
 "Problem: unnoticed temperature excursions during outages.",
 "Solution: battery-backed logger with RTC, alarm and exposure timer.",
 "Unique element: it reports total minutes above 8 C so the user can judge whether stock is still safe.",
 "Includes an outage detector so cuts are time-stamped."],
"principle": [
 "1. A DS18B20 inside the fridge reports temperature every 10 seconds.",
 "2. A DS3231 RTC provides time that survives power loss.",
 "3. If temperature is above 8 C or below 2 C, an excursion timer starts and a buzzer and LED warn.",
 "4. A transistor stage senses whether mains adapter voltage is present. Loss of mains is logged as an outage start time.",
 "5. The Nano runs from a 9V or Li-ion battery through a regulator during outages.",
 "6. Log data (time, temperature, status) is written to a microSD card each minute.",
 "7. The LCD shows current temperature, minimum, maximum and total excursion time."],
"sim": ["Wokwi: DS18B20 and DS3231 are both supported.",
        "Slide temperature above 8 C and confirm the timer and alarm.",
        "Toggle a switch for 'mains lost' and verify the log entry."],
"nets": [
 ("DS18B20 data","Nano D2 + 4.7k","Temperature"),
 ("DS3231 SDA, SCL","Nano A4, A5","Clock"),
 ("SD module MOSI, MISO, SCK, CS","Nano D11, D12, D13, D10","Logging"),
 ("Adapter sense divider","Nano A0","Mains present"),
 ("LCD I2C","Nano A4, A5","Display"),
 ("Nano D6","Buzzer","Alarm"),
 ("Nano D7, D8","Red and green LEDs","Status"),
 ("Battery + 1N4007 OR-ing","LM7805 to Nano","Backup power")],
"beme": [NANO,
 ("DS18B20 waterproof probe","",1,2500),("DS3231 RTC","With battery",1,3500),
 ("microSD module and card","Logging",1,3500),("16x2 I2C LCD","",1,3500),
 ("Buzzer","",1,250),("LM7805","",1,200),("9V battery and clip","",1,800),
 ("Enclosure","",1,2500)] + COMMON
},
{
"slug": "PowerWatch",
"title": "PowerWatch: Mains Outage and Voltage Quality Logger with Time Stamps",
"intro": [
 "Power supply in many Nigerian homes and hostels fluctuates and drops out many times a day, but nobody records how often or for how long.",
 "PowerWatch records every outage, its duration, and the minimum and maximum voltage seen, then shows a daily summary."],
"summary": [
 "Problem: no data on supply quality.",
 "Solution: voltage sensor, RTC, EEPROM log, LCD report and protective relay.",
 "Unique element: a daily 'supply score' (percentage of time with good voltage).",
 "Optional load protection through a relay with restore delay."],
"principle": [
 "1. A ZMPT101B module outputs an AC signal proportional to mains. The Nano samples it and computes RMS voltage.",
 "2. Voltage under 170 V or over 250 V counts as bad supply. Zero counts as an outage.",
 "3. A DS3231 RTC time-stamps the start and end of each event.",
 "4. Events are stored in EEPROM, and a backup battery keeps the Nano running during outages.",
 "5. The LCD cycles through current voltage, today's outage count, total outage time and supply score.",
 "6. The relay disconnects the load on bad voltage and reconnects after a 3 minute delay."],
"sim": ["Proteus or Wokwi: use a potentiometer in place of ZMPT101B and treat it as RMS voltage.",
        "Test bad-voltage events and confirm logs and the reconnect delay."],
"nets": [
 ("ZMPT101B OUT","Nano A0","Voltage sense"),
 ("DS3231 SDA, SCL","Nano A4, A5","Clock"),
 ("LCD I2C","Nano A4, A5","Display"),
 ("Nano D7","BC547 to relay","Load control"),
 ("Nano D3","Mode push button","Switch screens"),
 ("Nano D8","Buzzer","Alert"),
 ("Battery and 1N4007","LM7805 to Nano","Backup")],
"beme": [NANO,
 ("ZMPT101B module","",1,3500),("DS3231 RTC","",1,3500),("16x2 I2C LCD","",1,3500),
 ("5V relay module","10A",1,1500),("Buzzer","",1,250),("LM7805","",1,200),
 ("9V battery and clip","",1,800),("Enclosure and socket","",1,3500)] + COMMON
},
{
"slug": "RainRescue",
"title": "RainRescue: Automatic Rain-Sensing Clothes Line Retractor",
"intro": [
 "Students and families hang clothes outside and get caught by sudden rain. Someone has to be home to bring them in.",
 "RainRescue senses rain and darkness and drives a motor that pulls the clothes line under cover, then extends it again when it is dry and bright."],
"summary": [
 "Problem: wet laundry from sudden rain.",
 "Solution: rain sensor, LDR, DC motor with limit switches on an Arduino Nano.",
 "Unique element: it waits 10 minutes of dry weather before extending, to avoid repeated cycling.",
 "Uses transistors, diodes, an H-bridge and a relay-free motor drive."],
"principle": [
 "1. A rain sensor board gives an analog reading that drops when wet.",
 "2. An LDR in a divider reads light level.",
 "3. If the rain reading passes the limit, the Nano turns on an L298N or L293D to retract the motor.",
 "4. A limit switch at each end stops the motor.",
 "5. If the line is retracted, it extends only after 10 minutes of dry and bright readings.",
 "6. A manual switch overrides the automatic mode, and LEDs show the state."],
"sim": ["Wokwi or Proteus: potentiometers for rain and light.",
        "DC motor with H-bridge, push buttons as limit switches.",
        "Check the 10 minute delay by speeding it up in the code."],
"nets": [
 ("Rain sensor AO","Nano A0","Wet reading"),
 ("LDR + 10k divider","Nano A1","Light level"),
 ("Nano D5, D6, D9","L293D IN1, IN2, EN","Motor control"),
 ("Limit switch OUT","Nano D2","Retracted"),
 ("Limit switch IN","Nano D3","Extended"),
 ("Nano D7","Manual override switch","Mode"),
 ("Nano D10, D11","Status LEDs","Display"),
 ("9V to 12V supply","Motor, LM7805 to Nano","Power")],
"beme": [NANO,
 ("Rain sensor module","",1,1500),("LDR","",1,100),("L293D or L298N driver","",1,2000),
 ("DC gear motor","12V",1,3500),("Limit switches","",2,300),("LM7805","",1,200),
 ("Pulley, string, frame","",1,3500),("12V adapter","",1,2500)] + COMMON
}
]

def para(doc, text, size=11, bold=False, align=None, space=6):
    p = doc.add_paragraph()
    r = p.add_run(text); r.font.size = Pt(size); r.bold = bold
    p.paragraph_format.space_after = Pt(space)
    if align is not None: p.alignment = align

def table(doc, header, rows, size=9):
    t = doc.add_table(rows=1, cols=len(header)); t.style = "Table Grid"
    for i, h in enumerate(header):
        c = t.rows[0].cells[i]; c.text = ""
        c.paragraphs[0].add_run(h).bold = True
    for row in rows:
        cells = t.add_row().cells
        for i, v in enumerate(row):
            cells[i].text = ""
            cells[i].paragraphs[0].add_run(str(v)).font.size = Pt(size)

def build(n, p):
    d = Document()
    d.styles["Normal"].font.name = "Calibri"
    C = WD_ALIGN_PARAGRAPH.CENTER
    for _ in range(5): para(d, "")
    para(d, COURSE, 14, True, C)
    para(d, p["title"], 22, True, C, 24)
    para(d, f"Name: {NAME}", 14, False, C)
    para(d, f"Mat. No.: {MAT}", 14, False, C)
    para(d, DEPT, 12, False, C)
    d.add_page_break()
    d.add_heading("1. Introduction and Summary", 1)
    for t in p["intro"]: para(d, t)
    d.add_heading("Summary", 2)
    for s in p["summary"]: d.add_paragraph(s, style="List Bullet")
    d.add_heading("Objectives", 2)
    for s in ["Design the circuit in KiCad.", "Simulate in Proteus or Wokwi before building.",
              "Build, test and document results."]:
        d.add_paragraph(s, style="List Bullet")
    d.add_page_break()
    d.add_heading("2. Working Principle", 1)
    for t in p["principle"]: para(d, t)
    d.add_heading("Simulation plan", 2)
    for s in p["sim"]: d.add_paragraph(s, style="List Bullet")
    d.add_page_break()
    d.add_heading("3. Circuit Diagram (KiCad)", 1)
    img = f"kicad_{n}.png"
    if os.path.exists(img): d.add_picture(img, width=Inches(6))
    else: para(d, f"[Export your KiCad schematic as '{img}' next to this script and re-run.]")
    d.add_heading("Key connections", 2)
    table(d, ["From", "To", "Purpose"], p["nets"], 8)
    d.add_page_break()
    d.add_heading("4. BEME (Bill of Engineering Measurement and Evaluation)", 1)
    rows, total = [], 0
    for i, (item, spec, q, u) in enumerate(p["beme"], 1):
        amt = q * u; total += amt
        rows.append((i, item, spec, q, f"{u:,}", f"{amt:,}"))
    rows.append(("", "TOTAL", "", "", "", f"{total:,}"))
    table(d, ["S/N", "Item", "Specification", "Qty", "Unit (N)", "Amount (N)"], rows)
    para(d, "Prices are estimates. Confirm at local suppliers.", 9)
    name = f"Project{n}_{p['slug']}.docx"
    d.save(name); print("Saved", name, "Total N", f"{total:,}")

for i, p in enumerate(PROJECTS, 1):
    build(i, p)
