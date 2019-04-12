
from kivy.app import App
from kivy.uix.button import  Button
from kivy.uix.textinput import TextInput
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.tabbedpanel import TabbedPanel
from kivy.uix.label import Label
import excel_filler
from num2words import num2words

class BankFormFillApp(App):
    def build(self):
        self.mainBox = BoxLayout(orientation='horizontal', spacing=20)
        self.tab  = TabbedPanel()
        self.box = BoxLayout(orientation='vertical', spacing=20)
        self.payee = TextInput(hint_text='Payee', size_hint=(.7,.6))
        self.acc_num = TextInput(hint_text='Account Number ', size_hint=(.7,.6))
        self.ifsc = TextInput(hint_text='IFSC', size_hint=(.7,.6))
        self.bank = TextInput(hint_text='Bank', size_hint=(.7,.6))
        self.branch = TextInput(hint_text='Bank Branch', size_hint=(.7,.6))
        self.amount = TextInput(hint_text='Amount', size_hint=(.7,.6))
        amount_words = num2words(int(self.amount._get_text()))
        self.label = Label(text='Amount in words:')
        self.amount_in_words = Label(text=amount_words)
        self.date = TextInput(hint_text='Date', size_hint=(.7,.6))

        self.save = Button(text='Save Data', on_press=self.save_data, size_hint=(.7,.6))

        self.box.add_widget(self.payee)
        self.box.add_widget(self.acc_num)
        self.box.add_widget(self.ifsc)
        self.box.add_widget(self.bank)
        self.box.add_widget(self.branch)
        self.box.add_widget(self.amount)
        self.box.add_widget(self.amount_words)
        self.box.add_widget(self.date)

        self.box.add_widget(self.save)

        self.box1 = BoxLayout(orientation='vertical', spacing = 20)
        self.search_record_by_name = TextInput(hint_text='Enter name to search And create NEFT/RTGS Form', size_hint=(.5,.2))
        self.create = Button(text='Create Neft/RTGS', on_press=self.create_form, size_hint=(.5,.2))

        self.box1.add_widget(self.search_record_by_name)
        self.box1.add_widget(self.create)

        self.mainBox.add_widget(self.box)
        self.mainBox.add_widget(self.box1)
        return self.mainBox

    def save_data(self, instance):
        excel_filler.add_entry(self.payee, self.acc_num, self.ifsc, self.bank, self.branch, self.amount, self.amount_words, self.date)
        excel_filler.add_to_json(self.payee, self.acc_num, self.ifsc, self.bank, self.branch, self.amount, self.amount_words, self.date)

    def create_form(self,instance):
        excel_filler.search_entry_and_generate_doc(self.search_record_by_name)

BankFormFillApp().run()