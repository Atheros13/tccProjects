### PRIME IMPORTS ###

from __future__ import division

### KIVY IMPORTS ###

from kivy.app import App
from kivy.clock import Clock
from kivy.core.window import Window
from kivy.graphics import Color, Rectangle
from kivy.graphics.vertex_instructions import Line
from kivy.lang import Builder
from kivy.properties import StringProperty, ObjectProperty, NumericProperty
from kivy.properties import ListProperty, DictProperty, BooleanProperty

from kivy.uix.button import Button
from kivy.uix.carousel import Carousel
from kivy.uix.checkbox import CheckBox
from kivy.uix.colorpicker import ColorPicker, ColorWheel
from kivy.uix.dropdown import DropDown
from kivy.uix.filechooser import FileChooserIconView, FileChooserListView
from kivy.uix.image import Image, AsyncImage
from kivy.uix.label import Label 
from kivy.uix.popup import Popup
from kivy.uix.textinput import TextInput

from kivy.uix.boxlayout import BoxLayout
from kivy.uix.gridlayout import GridLayout
from kivy.uix.scrollview import ScrollView
from kivy.uix.tabbedpanel import TabbedPanel, TabbedPanelItem

class QuickSignGUI(App):

	version = StringProperty('Login') #
	title = 'TCC Quick Login'

	def __init__(self, *args, **kwargs):
		super(QuickSignGUI, self).__init__(**kwargs)

		## SETTINGS
		Window.size = (200, 100)

		## DATABASE


	def build(self):
		
		gui = BoxLayout(orientation="horizontal")

		gui.add_widget(self.build_name())
		gui.add_widget(self.build_role())
		gui.add_widget(self.build_button())

		return gui

	def build_name(self):

		pass

	def build_role(self):

		pass

	def build_button(self):

		pass