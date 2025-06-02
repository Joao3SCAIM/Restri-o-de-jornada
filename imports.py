from selenium.common.exceptions import NoSuchElementException, TimeoutException
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.edge.service import Service
from selenium.webdriver.common.keys import Keys
from pathlib import Path
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font
from dotenv import load_dotenv
import schedule
import subprocess
import time
import datetime
import os
