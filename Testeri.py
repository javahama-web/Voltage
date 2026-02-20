import pyvisa, time
import pandas as pd
import random  # Tuodaan kirjasto käyttöön
from datetime import datetime
from abc import ABC, abstractmethod


# --- RAJAPINNAT (Abstract Base Classes) ---

class MeasurementDevice(ABC):
    """Abstrakti luokka, joka määrittelee mittalaitteen rajapinnan."""
    
    @abstractmethod
    def connect(self, address):
        pass

    @abstractmethod
    def get_reading(self):
        pass

    @abstractmethod
    def disconnect(self):
        pass

# --- LAITETOTEUTUKSET ---

class MultimeterAgilent(MeasurementDevice):
    """Toteutus Agilent/Keysight yleismittarille."""
    """Simuloitu mittalaite testausta varten lisäämäll ('@py'."""
    
    def __init__(self):
        self.rm = pyvisa.ResourceManager('@py')
        self.instrument = None

    def connect(self, address):
        # Esim. address = "GPIB0::7::INSTR"
        #########self.instrument = self.rm.open_resource(address)
        #########print(f"Yhdistetty laitteeseen: {self.instrument.query('*IDN?')}")
        print(f"SIMULAATIO: Yhdistetty osoitteeseen {address}")
        
    def get_reading(self):
        # SCPI-komento jännitteen mittaukseen
        #########result = self.instrument.query("MEASure:VOLTage:DC?")
        ########return float(result)
        val = random.uniform(4.95, 5.05)
        return val

    
    def disconnect(self):
        print("SIMULAATIO: Yhteys katkaistu.")
        ########if self.instrument:
            #########self.instrument.close()

# --- DATAN HALLINTA JA EXCEL ---

class ResultsManager:
    """Hoitaa tulosten keräämisen ja tallentamisen Exceliin."""
    
    def __init__(self):
        self.data = []

    def add_result(self, test_point, value, unit):
        self.data.append({
            "Timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "Test Point": test_point,
            "Value": value,
            "Unit": unit
        })

    def save_to_excel(self, filename="mittaustulokset.xlsx"):
        df = pd.DataFrame(self.data)
        # Käytetään pandasia Excel-vientiin
        df.to_excel(filename, index=False, engine='openpyxl')
        print(f"Tulokset tallennettu tiedostoon: {filename}")

# --- TESTIMOOTTORI (Orchestrator) ---

class TestController:
    """Ohjaa testauksen kulkua."""
    
    def __init__(self, device: MeasurementDevice, results_manager: ResultsManager):
        self.device = device
        self.results = results_manager

    def run_sequence(self, visa_address):
        try:
            self.device.connect(visa_address)
            
            # Suoritetaan testivaiheet
            print("Aloitetaan mittaus...")
            val = self.device.get_reading()
            self.results.add_result("VCC_Input", val, "V")
            
            # Lisää tähän muita vaiheita...
            
        except Exception as e:
            print(f"Virhe testin aikana: {e}")
        finally:
            self.device.disconnect()
            self.results.save_to_excel()

# --- PÄÄOHJELMA ---

if __name__ == "__main__":
    # 1. Alustetaan komponentit
    mittari = MultimeterAgilent()
    tallennin = ResultsManager()
    i=0
    
    # 2. Luodaan ohjain ja suoritetaan testi
    tester = TestController(mittari, tallennin)

    for i in range(50):
        tester.run_sequence("USB0::0x0957::0x0607::MY53000362::INSTR")
        time.sleep(5)








