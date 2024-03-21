import os
import subprocess
import tempfile
from PyQt6 import QtCore, QtGui, QtWidgets
import docx
from PyQt6 import *
from PyQt6.QtCore import *
from PyQt6.QtGui import *
from PyQt6.QtWidgets import *
from colorama import win32
from docx2pdf import convert
from docxtpl import DocxTemplate
from docxtpl import *
import datetime
import sys



try:
    class Moulin_Window():
        ##############################################buttuons&def####################
        def psbd_calcull(self):
            self.rpsbd.clear()
            self.bpsbd.clear()
            while self.vpsbd.text() >= '72':
                self.rpsbd.setValue(10.40)
                if self.vpsbd.text() <= '72.24':
                    self.rpsbd.setValue(10.40)
                elif self.vpsbd.text() >= '72.26':
                    self.rpsbd.setValue(9.90)
                    if self.vpsbd.text() <= '72.49':
                        self.rpsbd.setValue(9.90)
                    elif self.vpsbd.text() >= '72.50':
                        self.rpsbd.setValue(9.40)
                        if self.vpsbd.text() <= '72.74':
                            self.rpsbd.setValue(9.40)
                        elif self.vpsbd.text() >= '72.75':
                            self.rpsbd.setValue(8.90)
                            if self.vpsbd.text() <= '72.99':
                                self.rpsbd.setValue(8.90)
                            elif self.vpsbd.text() >= '73':
                                self.rpsbd.setValue(8.40)
                                if self.vpsbd.text() <= '73.24':
                                    self.rpsbd.setValue(8.40)
                                elif self.vpsbd.text() >= '73.25':
                                    self.rpsbd.setValue(7.90)
                                    if self.vpsbd.text() <= '73.49':
                                        self.rpsbd.setValue(7.90)
                                    elif self.vpsbd.text() >= '73.50':
                                        self.rpsbd.setValue(7.40)
                                        if self.vpsbd.text() <= '73.74':
                                            self.rpsbd.setValue(7.40)
                                        elif self.vpsbd.text() >= '73.75':
                                            self.rpsbd.setValue(6.90)
                                            if self.vpsbd.text() <= '73.99':
                                                self.rpsbd.setValue(6.90)
                                            elif self.vpsbd.text() >= '74':
                                                self.rpsbd.setValue(6.40)
                                                if self.vpsbd.text() <= '74.24':
                                                    self.rpsbd.setValue(6.40)
                                                elif self.vpsbd.text() >= '74.25':
                                                    self.rpsbd.setValue(5.90)
                                                    if self.vpsbd.text() <= '74.49':
                                                        self.rpsbd.setValue(5.90)
                                                    elif self.vpsbd.text() >= '74.50':
                                                        self.rpsbd.setValue(5.40)
                                                        if self.vpsbd.text() <= '74.74':
                                                            self.rpsbd.setValue(5.40)
                                                        elif self.vpsbd.text() >= '74.75':
                                                            self.rpsbd.setValue(4.90)
                                                            if self.vpsbd.text() <= '74.99':
                                                                self.rpsbd.setValue(4.90)
                                                            elif self.vpsbd.text() >= '75':
                                                                self.rpsbd.setValue(4.40)
                                                                if self.vpsbd.text() <= '75.24':
                                                                    self.rpsbd.setValue(4.40)
                                                                elif self.vpsbd.text() >= '75.25':
                                                                    self.rpsbd.setValue(3.90)
                                                                    if self.vpsbd.text() <= '75.49':
                                                                        self.rpsbd.setValue(3.90)
                                                                    elif self.vpsbd.text() >= '75.50':
                                                                        self.rpsbd.setValue(3.40)
                                                                        if self.vpsbd.text() <= '75.74':
                                                                            self.rpsbd.setValue(3.40)
                                                                        elif self.vpsbd.text() >= '75.75':
                                                                            self.rpsbd.setValue(2.90)
                                                                            if self.vpsbd.text() <= '75.99':
                                                                                self.rpsbd.setValue(2.90)
                                                                            elif self.vpsbd.text() >= '76':
                                                                                self.rpsbd.setValue(2.40)
                                                                                if self.vpsbd.text() <= '76.24':
                                                                                    self.rpsbd.setValue(2.40)
                                                                                elif self.vpsbd.text() >= '76.25':
                                                                                    self.rpsbd.setValue(2.05)
                                                                                    if self.vpsbd.text() <= '76.49':
                                                                                        self.rpsbd.setValue(2.05)
                                                                                    elif self.vpsbd.text() >= '76.50':
                                                                                        self.rpsbd.setValue(1.70)
                                                                                        if self.vpsbd.text() <= '76.74':
                                                                                            self.rpsbd.setValue(1.70)
                                                                                        elif self.vpsbd.text() >= '76.75':
                                                                                            self.rpsbd.setValue(1.35)
                                                                                            if self.vpsbd.text() <= '76.99':
                                                                                                self.rpsbd.setValue(
                                                                                                    1.35)
                                                                                            elif self.vpsbd.text() >= '77':
                                                                                                self.rpsbd.setValue(
                                                                                                    1.00)
                                                                                                if self.vpsbd.text() <= '77.24':
                                                                                                    self.rpsbd.setValue(
                                                                                                        1.00)
                                                                                                elif self.vpsbd.text() >= '77.25':
                                                                                                    self.rpsbd.setValue(
                                                                                                        0.75)
                                                                                                    if self.vpsbd.text() <= '77.49':
                                                                                                        self.rpsbd.setValue(
                                                                                                            0.75)
                                                                                                    elif self.vpsbd.text() >= '77.50':
                                                                                                        self.rpsbd.setValue(
                                                                                                            0.50)
                                                                                                        if self.vpsbd.text() <= '77.74':
                                                                                                            self.rpsbd.setValue(
                                                                                                                0.50)
                                                                                                        elif self.vpsbd.text() >= '77.75':
                                                                                                            self.rpsbd.setValue(
                                                                                                                0.25)
                                                                                                            if self.vpsbd.text() <= '77.99':
                                                                                                                self.rpsbd.setValue(
                                                                                                                    0.25)
                                                                                                            else:
                                                                                                                self.rpsbd.clear()
                                                                                                                self.rpsbd.setValue(
                                                                                                                    0)
                break

            while self.vpsbd.text() >= '78':
                self.bpsbd.setValue(0.15)
                if self.vpsbd.text() >= '78':
                    self.bpsbd.setValue(0.15)
                    if self.vpsbd.text() <= '78.25':
                        self.bpsbd.setValue(0.15)
                    elif self.vpsbd.text() >= '78.26':
                        self.bpsbd.setValue(0.30)
                        if self.vpsbd.text() <= '778.50':
                            self.bpsbd.setValue(0.30)
                        elif self.vpsbd.text() >= '78.51':
                            self.bpsbd.setValue(0.45)
                            if self.vpsbd.text() <= '78.75':
                                self.bpsbd.setValue(0.45)
                            elif self.vpsbd.text() >= '78.76':
                                self.bpsbd.setValue(0.60)
                                if self.vpsbd.text() <= '79':
                                    self.bpsbd.setValue(0.60)
                                elif self.vpsbd.text() >= '79.01':
                                    self.bpsbd.setValue(0.75)
                                    if self.vpsbd.text() <= '79.25':
                                        self.bpsbd.setValue(0.75)
                                    elif self.vpsbd.text() >= '79.26':
                                        self.bpsbd.setValue(0.90)
                                        if self.vpsbd.text() <= '79.50':
                                            self.bpsbd.setValue(0.90)
                                        elif self.vpsbd.text() >= '79.51':
                                            self.bpsbd.setValue(1.05)
                                            if self.vpsbd.text() <= '79.75':
                                                self.bpsbd.setValue(1.05)
                                            elif self.vpsbd.text() >= '79.76':
                                                self.bpsbd.setValue(1.20)
                                                if self.vpsbd.text() <= '80':
                                                    self.bpsbd.setValue(1.20)
                                                elif self.vpsbd.text() >= '80.01':
                                                    self.bpsbd.setValue(1.35)
                                                    if self.vpsbd.text() <= '80.25':
                                                        self.bpsbd.setValue(1.35)
                                                    elif self.vpsbd.text() >= '80.25':
                                                        self.bpsbd.setValue(1.50)
                                                        if self.vpsbd.text() <= '80.50':
                                                            self.bpsbd.setValue(1.50)
                                                        elif self.vpsbd.text() >= '80.51':
                                                            self.bpsbd.setValue(1.65)
                                                            if self.vpsbd.text() <= '80.75':
                                                                self.bpsbd.setValue(1.65)
                                                            elif self.vpsbd.text() >= '80.76':
                                                                self.bpsbd.setValue(1.80)
                                                                if self.vpsbd.text() <= '81':
                                                                    self.bpsbd.setValue(1.80)
                                                                elif self.vpsbd.text() >= '81.01':
                                                                    self.bpsbd.setValue(1.95)
                                                                    if self.vpsbd.text() <= '81.25':
                                                                        self.bpsbd.setValue(1.95)
                                                                    elif self.vpsbd.text() >= '81.26':
                                                                        self.bpsbd.setValue(2.10)
                                                                        if self.vpsbd.text() <= '81.50':
                                                                            self.bpsbd.setValue(2.10)
                                                                        elif self.vpsbd.text() >= '81.51':
                                                                            self.bpsbd.setValue(2.25)
                                                                            if self.vpsbd.text() <= '81.75':
                                                                                self.bpsbd.setValue(2.25)
                                                                            elif self.vpsbd.text() >= '81.76':
                                                                                self.bpsbd.setValue(2.40)
                                                                                if self.vpsbd.text() <= '82':
                                                                                    self.bpsbd.setValue(2.40)
                                                                                elif self.vpsbd.text() >= '82.01':
                                                                                    self.bpsbd.setValue(2.50)
                                                                                    if self.vpsbd.text() <= '82.25':
                                                                                        self.bpsbd.setValue(2.50)
                                                                                    elif self.vpsbd.text() >= '82.26':
                                                                                        self.bpsbd.setValue(2.60)
                                                                                        if self.vpsbd.text() <= '82.50':
                                                                                            self.bpsbd.setValue(2.60)
                                                                                        elif self.vpsbd.text() >= '82.51':
                                                                                            self.bpsbd.setValue(2.70)
                                                                                            if self.vpsbd.text() <= '82.75':
                                                                                                self.bpsbd.setValue(
                                                                                                    2.70)
                                                                                            elif self.vpsbd.text() >= '82.76':
                                                                                                self.bpsbd.setValue(
                                                                                                    2.80)
                                                                                                if self.vpsbd.text() <= '83':
                                                                                                    self.bpsbd.setValue(
                                                                                                        2.80)
                                                                                                elif self.vpsbd.text() >= '83.01':
                                                                                                    self.bpsbd.setValue(
                                                                                                        2.85)
                                                                                                    if self.vpsbd.text() <= '83.25':
                                                                                                        self.bpsbd.setValue(
                                                                                                            2.85)
                                                                                                    elif self.vpsbd.text() >= '83.26':
                                                                                                        self.bpsbd.setValue(
                                                                                                            2.90)
                                                                                                        if self.vpsbd.text() <= '83.50':
                                                                                                            self.bpsbd.setValue(
                                                                                                                2.90)
                                                                                                        elif self.vpsbd.text() >= '83.51':
                                                                                                            self.bpsbd.setValue(
                                                                                                                2.95)
                                                                                                            if self.vpsbd.text() <= '83.75':
                                                                                                                self.bpsbd.setValue(
                                                                                                                    2.95)
                                                                                                            elif self.vpsbd.text() >= '83.76':
                                                                                                                self.bpsbd.setValue(
                                                                                                                    3)
                                                                                                                if self.vpsbd.text() <= '84':
                                                                                                                    self.bpsbd.setValue(
                                                                                                                        3)

                                                                                                                else:
                                                                                                                    self.bpsbd.clear()
                                                                                                                    self.bpsbd.setValue(
                                                                                                                        0)
                break

        def imp1erbd(self):
            while self.vtotalprembd.value() >= 0.10:
                self.btotalprembd.setValue(0.60)
                if self.vtotalprembd.value() <= 0.25:
                    self.btotalprembd.setValue(0.60)
                elif self.vtotalprembd.value() >= 0.26 and self.vtotalprembd.value() <= 0.50:
                    self.btotalprembd.setValue(0.45)
                elif self.vtotalprembd.value() >= 0.51 and self.vtotalprembd.value() <= 0.75:
                    self.btotalprembd.setValue(0.30)
                elif self.vtotalprembd.value() >= 0.76 and self.vtotalprembd.value() <= 1.00:
                    self.btotalprembd.setValue(0.15)
                else:
                    self.btotalprembd.clear()
                    self.btotalprembd.setValue(0)
                break
            while self.vtotalprembd.value() >= 1.01:
                self.rtotalprembd.setValue(0.15)
                if self.vtotalprembd.value() <= 1.25:
                    self.rtotalprembd.setValue(0.15)
                elif self.vtotalprembd.value() >= 1.26 and self.vtotalprembd.value() <= 1.50:
                    self.rtotalprembd.setValue(0.30)
                elif self.vtotalprembd.value() >= 1.51 and self.vtotalprembd.value() <= 1.75:
                    self.rtotalprembd.setValue(0.45)
                elif self.vtotalprembd.value() >= 1.76 and self.vtotalprembd.value() <= 2.00:
                    self.rtotalprembd.setValue(0.60)
                elif self.vtotalprembd.value() >= 2.01 and self.vtotalprembd.value() <= 2.25:
                    self.rtotalprembd.setValue(0.75)
                elif self.vtotalprembd.value() >= 2.26 and self.vtotalprembd.value() <= 2.50:
                    self.rtotalprembd.setValue(0.90)
                elif self.vtotalprembd.value() >= 2.51 and self.vtotalprembd.value() <= 2.75:
                    self.rtotalprembd.setValue(1.05)
                elif self.vtotalprembd.value() >= 2.76 and self.vtotalprembd.value() <= 3.00:
                    self.rtotalprembd.setValue(1.20)
                elif self.vtotalprembd.value() >= 3.01 and self.vtotalprembd.value() <= 3.25:
                    self.rtotalprembd.setValue(1.35)
                elif self.vtotalprembd.value() >= 3.26 and self.vtotalprembd.value() <= 3.50:
                    self.rtotalprembd.setValue(1.50)
                elif self.vtotalprembd.value() >= 3.51 and self.vtotalprembd.value() <= 3.75:
                    self.rtotalprembd.setValue(1.65)
                elif self.vtotalprembd.value() >= 3.76 and self.vtotalprembd.value() <= 4.00:
                    self.rtotalprembd.setValue(1.80)
                elif self.vtotalprembd.value() >= 4.01 and self.vtotalprembd.value() <= 4.25:
                    self.rtotalprembd.setValue(1.95)
                elif self.vtotalprembd.value() >= 4.26 and self.vtotalprembd.value() <= 4.50:
                    self.rtotalprembd.setValue(2.10)
                elif self.vtotalprembd.value() >= 4.51 and self.vtotalprembd.value() <= 4.75:
                    self.rtotalprembd.setValue(2.25)
                elif self.vtotalprembd.value() >= 4.76 and self.vtotalprembd.value() <= 5.00:
                    self.rtotalprembd.setValue(2.40)
                elif self.vtotalprembd.value() >= 5.01 and self.vtotalprembd.value() <= 5.25:
                    self.rtotalprembd.setValue(2.55)
                elif self.vtotalprembd.value() >= 5.26 and self.vtotalprembd.value() <= 5.50:
                    self.rtotalprembd.setValue(2.70)
                elif self.vtotalprembd.value() >= 5.51 and self.vtotalprembd.value() <= 5.75:
                    self.rtotalprembd.setValue(2.85)
                elif self.vtotalprembd.value() >= 5.76 and self.vtotalprembd.value() <= 6.00:
                    self.rtotalprembd.setValue(3.00)
                elif self.vtotalprembd.value() >= 6.01 and self.vtotalprembd.value() <= 6.25:
                    self.rtotalprembd.setValue(3.15)
                elif self.vtotalprembd.value() >= 6.26 and self.vtotalprembd.value() <= 6.50:
                    self.rtotalprembd.setValue(3.30)
                elif self.vtotalprembd.value() >= 6.51 and self.vtotalprembd.value() <= 6.75:
                    self.rtotalprembd.setValue(3.45)
                elif self.vtotalprembd.value() >= 6.76 and self.vtotalprembd.value() <= 7.00:
                    self.rtotalprembd.setValue(3.60)
                elif self.vtotalprembd.value() >= 7.01 and self.vtotalprembd.value() <= 7.25:
                    self.rtotalprembd.setValue(3.75)
                elif self.vtotalprembd.value() >= 7.26 and self.vtotalprembd.value() <= 7.50:
                    self.rtotalprembd.setValue(3.90)
                elif self.vtotalprembd.value() >= 7.51 and self.vtotalprembd.value() <= 7.75:
                    self.rtotalprembd.setValue(4.05)
                elif self.vtotalprembd.value() >= 7.76 and self.vtotalprembd.value() <= 8.00:
                    self.rtotalprembd.setValue(4.20)
                elif self.vtotalprembd.value() >= 8.01 and self.vtotalprembd.value() <= 8.25:
                    self.rtotalprembd.setValue(4.35)
                elif self.vtotalprembd.value() >= 8.26 and self.vtotalprembd.value() <= 8.50:
                    self.rtotalprembd.setValue(4.50)
                elif self.vtotalprembd.value() >= 8.51 and self.vtotalprembd.value() <= 8.75:
                    self.rtotalprembd.setValue(4.65)
                elif self.vtotalprembd.value() >= 8.76 and self.vtotalprembd.value() <= 9.00:
                    self.rtotalprembd.setValue(4.80)
                else:
                    self.rtotalprembd.clear()
                    self.rtotalprembd.setValue(0)
                break

        def imp2embd(self):
            while self.vtotaldembd.value() >= 10.01:
                self.rtotaldembd.setValue(0.075)
                if self.vtotaldembd.value() <= 10.25:
                    self.rtotaldembd.setValue(0.075)
                elif self.vtotaldembd.value() >= 10.26 and self.vtotaldembd.value() <= 10.50:
                    self.rtotaldembd.setValue(0.15)
                elif self.vtotaldembd.value() >= 10.51 and self.vtotaldembd.value() <= 10.75:
                    self.rtotaldembd.setValue(0.225)
                elif self.vtotaldembd.value() >= 10.76 and self.vtotaldembd.value() <= 11.00:
                    self.rtotaldembd.setValue(0.30)
                elif self.vtotaldembd.value() >= 11.01 and self.vtotaldembd.value() <= 11.25:
                    self.rtotaldembd.setValue(0.375)
                elif self.vtotaldembd.value() >= 11.26 and self.vtotaldembd.value() <= 11.50:
                    self.rtotaldembd.setValue(0.45)
                elif self.vtotaldembd.value() >= 11.51 and self.vtotaldembd.value() <= 11.75:
                    self.rtotaldembd.setValue(0.525)
                elif self.vtotaldembd.value() >= 11.76 and self.vtotaldembd.value() <= 12.00:
                    self.rtotaldembd.setValue(0.60)
                elif self.vtotaldembd.value() >= 12.01 and self.vtotaldembd.value() <= 12.25:
                    self.rtotaldembd.setValue(0.675)
                elif self.vtotaldembd.value() >= 12.26 and self.vtotaldembd.value() <= 12.50:
                    self.rtotaldembd.setValue(0.75)
                elif self.vtotaldembd.value() >= 12.51 and self.vtotaldembd.value() <= 12.75:
                    self.rtotaldembd.setValue(0.825)
                elif self.vtotaldembd.value() >= 12.76 and self.vtotaldembd.value() <= 13.00:
                    self.rtotaldembd.setValue(0.90)
                elif self.vtotaldembd.value() >= 13.01 and self.vtotaldembd.value() <= 13.25:
                    self.rtotaldembd.setValue(0.975)
                elif self.vtotaldembd.value() >= 13.26 and self.vtotaldembd.value() <= 13.50:
                    self.rtotaldembd.setValue(1.05)
                elif self.vtotaldembd.value() >= 13.51 and self.vtotaldembd.value() <= 13.75:
                    self.rtotaldembd.setValue(1.125)
                elif self.vtotaldembd.value() >= 13.76 and self.vtotaldembd.value() <= 14.00:
                    self.rtotaldembd.setValue(1.20)
                elif self.vtotaldembd.value() >= 14.01 and self.vtotaldembd.value() <= 14.25:
                    self.rtotaldembd.setValue(1.275)
                elif self.vtotaldembd.value() >= 14.26 and self.vtotaldembd.value() <= 14.50:
                    self.rtotaldembd.setValue(1.35)
                elif self.vtotaldembd.value() >= 14.51 and self.vtotaldembd.value() <= 14.75:
                    self.rtotaldembd.setValue(1.425)
                elif self.vtotaldembd.value() >= 14.76 and self.vtotaldembd.value() <= 15.00:
                    self.rtotaldembd.setValue(1.50)
                elif self.vtotaldembd.value() <= 15.01 and self.vtotaldembd.value() <= 15.25:
                    self.rtotaldembd.setValue(1.60)
                elif self.vtotaldembd.value() >= 15.26 and self.vtotaldembd.value() <= 15.50:
                    self.rtotaldembd.setValue(1.70)
                elif self.vtotaldembd.value() >= 15.51 and self.vtotaldembd.value() <= 15.75:
                    self.rtotaldembd.setValue(1.80)
                elif self.vtotaldembd.value() >= 15.76 and self.vtotaldembd.value() <= 16.00:
                    self.rtotaldembd.setValue(1.90)
                elif self.vtotaldembd.value() >= 16.01 and self.vtotaldembd.value() <= 16.25:
                    self.rtotaldembd.setValue(2.00)
                elif self.vtotaldembd.value() >= 16.26 and self.vtotaldembd.value() <= 16.50:
                    self.rtotaldembd.setValue(2.10)
                elif self.vtotaldembd.value() >= 16.51 and self.vtotaldembd.value() <= 16.75:
                    self.rtotaldembd.setValue(2.20)
                elif self.vtotaldembd.value() >= 11.76 and self.vtotaldembd.value() <= 17.00:
                    self.rtotaldembd.setValue(2.30)
                elif self.vtotaldembd.value() >= 17.01 and self.vtotaldembd.value() <= 17.25:
                    self.rtotaldembd.setValue(2.40)
                elif self.vtotaldembd.value() >= 17.26 and self.vtotaldembd.value() <= 17.50:
                    self.rtotaldembd.setValue(2.50)
                elif self.vtotaldembd.value() >= 17.51 and self.vtotaldembd.value() <= 17.75:
                    self.rtotaldembd.setValue(2.60)
                elif self.vtotaldembd.value() >= 17.76 and self.vtotaldembd.value() <= 18.00:
                    self.rtotaldembd.setValue(2.70)
                elif self.vtotaldembd.value() >= 18.01 and self.vtotaldembd.value() <= 18.25:
                    self.rtotaldembd.setValue(2.80)
                elif self.vtotaldembd.value() >= 18.26 and self.vtotaldembd.value() <= 18.50:
                    self.rtotaldembd.setValue(2.90)
                elif self.vtotaldembd.value() >= 18.51 and self.vtotaldembd.value() <= 18.75:
                    self.rtotaldembd.setValue(3.00)
                elif self.vtotaldembd.value() >= 18.76 and self.vtotaldembd.value() <= 19.00:
                    self.rtotaldembd.setValue(3.10)
                elif self.vtotaldembd.value() >= 19.01 and self.vtotaldembd.value() <= 19.25:
                    self.rtotaldembd.setValue(3.20)
                elif self.vtotaldembd.value() >= 19.26 and self.vtotaldembd.value() <= 19.50:
                    self.rtotaldembd.setValue(3.30)
                elif self.vtotaldembd.value() >= 19.51 and self.vtotaldembd.value() <= 19.75:
                    self.rtotaldembd.setValue(3.40)
                elif self.vtotaldembd.value() >= 19.76 and self.vtotaldembd.value() <= 20.00:
                    self.rtotaldembd.setValue(3.50)
                else:
                    self.rtotaldembd.setValue(0.0)
                break

        def ref_graincassébd(self):
            while self.vgraincassébd.value() >= 3.01:
                self.rgraincassébd.setValue(0.05)
                if self.vgraincassébd.value() <= 3.25:
                    self.rgraincassébd.setValue(0.05)
                elif self.vgraincassébd.value() >= 3.26 and self.vgraincassébd.value() <= 3.50:
                    self.rgraincassébd.setValue(0.10)
                elif self.vgraincassébd.value() >= 3.51 and self.vgraincassébd.value() <= 3.75:
                    self.rgraincassébd.setValue(0.15)
                elif self.vgraincassébd.value() >= 3.76 and self.vgraincassébd.value() <= 4.00:
                    self.rgraincassébd.setValue(0.20)
                elif self.vgraincassébd.value() >= 4.01 and self.vgraincassébd.value() <= 4.25:
                    self.rgraincassébd.setValue(0.25)
                elif self.vgraincassébd.value() >= 4.26 and self.vgraincassébd.value() <= 4.50:
                    self.rgraincassébd.setValue(0.30)
                elif self.vgraincassébd.value() >= 4.51 and self.vgraincassébd.value() <= 4.75:
                    self.rgraincassébd.setValue(0.35)
                elif self.vgraincassébd.value() >= 4.76 and self.vgraincassébd.value() <= 5.00:
                    self.rgraincassébd.setValue(0.40)
                elif self.vgraincassébd.value() >= 5.01 and self.vgraincassébd.value() <= 5.25:
                    self.rgraincassébd.setValue(0.47)
                elif self.vgraincassébd.value() >= 5.26 and self.vgraincassébd.value() <= 5.50:
                    self.rgraincassébd.setValue(0.55)
                elif self.vgraincassébd.value() >= 5.51 and self.vgraincassébd.value() <= 5.75:
                    self.rgraincassébd.setValue(0.62)
                elif self.vgraincassébd.value() >= 5.76 and self.vgraincassébd.value() <= 6.00:
                    self.rgraincassébd.setValue(0.70)
                elif self.vgraincassébd.value() >= 6.01 and self.vgraincassébd.value() <= 6.25:
                    self.rgraincassébd.setValue(0.77)
                elif self.vgraincassébd.value() >= 6.26 and self.vgraincassébd.value() <= 6.50:
                    self.rgraincassébd.setValue(0.85)
                elif self.vgraincassébd.value() >= 6.51 and self.vgraincassébd.value() <= 6.75:
                    self.rgraincassébd.setValue(0.92)
                elif self.vgraincassébd.value() >= 6.76 and self.vgraincassébd.value() <= 7.00:
                    self.rgraincassébd.setValue(1)
                else:
                    self.rgraincassébd.setValue(0)
                break

        def ref_grainbouté_bd(self):
            while self.vgrainboutébd.value() >= 4.01:
                self.rgrainboutébd.setValue(0.05)
                if self.vgrainboutébd.value() <= 5:
                    self.rgrainboutébd.setValue(0.05)
                elif self.vgrainboutébd.value() >= 5.01 and self.vgrainboutébd.value() <= 6:
                    self.rgrainboutébd.setValue(0.15)
                elif self.vgrainboutébd.value() >= 6.01 and self.vgrainboutébd.value() <= 7:
                    self.rgrainboutébd.setValue(0.25)
                elif self.vgrainboutébd.value() >= 7.01 and self.vgrainboutébd.value() <= 8:
                    self.rgrainboutébd.setValue(0.35)
                elif self.vgrainboutébd.value() >= 8.01 and self.vgrainboutébd.value() <= 9:
                    self.rgrainboutébd.setValue(0.45)
                elif self.vgrainboutébd.value() >= 9.01 and self.vgrainboutébd.value() <= 10:
                    self.rgrainboutébd.setValue(0.55)
                elif self.vgrainboutébd.value() >= 10.01 and self.vgrainboutébd.value() <= 11:
                    self.rgrainboutébd.setValue(0.65)
                elif self.vgrainboutébd.value() >= 11.01 and self.vgrainboutébd.value() <= 12:
                    self.rgrainboutébd.setValue(0.75)
                elif self.vgrainboutébd.value() >= 12.01 and self.vgrainboutébd.value() <= 13:
                    self.rgrainboutébd.setValue(0.85)
                elif self.vgrainboutébd.value() >= 13.01 and self.vgrainboutébd.value() <= 14:
                    self.rgrainboutébd.setValue(0.95)
                else:
                    self.rgrainboutébd.setVlue(0)
                break

        def boni_bltdbd(self):
            while self.vblétendreinbledur.value() >= 1:
                self.bblétendreinbledur.setValue(0.26)
                if self.vblétendreinbledur.value() <= 9:
                    self.bblétendreinbledur.setValue(0.26)
                elif self.vblétendreinbledur.value() >= 9.01 and self.vblétendreinbledur.value() <= 10:
                    self.bblétendreinbledur.setValue(0.195)
                elif self.vblétendreinbledur.value() >= 10.01 and self.vblétendreinbledur.value() <= 11:
                    self.bblétendreinbledur.setValue(0.130)
                elif self.vblétendreinbledur.value() >= 11.01 and self.vblétendreinbledur.value() <= 12:
                    self.bblétendreinbledur.setValue(0.065)
                else:
                    self.bblétendreinbledur.setValue(0)

                if self.vblétendreinbledur.value() >= 12.01 and self.vblétendreinbledur.value() <= 13:
                    self.rblétendreinbledur.setValue(0.065)
                elif self.vblétendreinbledur.value() >= 13.01 and self.vblétendreinbledur.value() <= 14:
                    self.rblétendreinbledur.setValue(0.140)
                elif self.vblétendreinbledur.value() >= 14.01 and self.vblétendreinbledur.value() <= 15:
                    self.rblétendreinbledur.setValue(0.225)
                elif self.vblétendreinbledur.value() >= 15.01 and self.vblétendreinbledur.value() <= 16:
                    self.rblétendreinbledur.setValue(0.320)
                elif self.vblétendreinbledur.value() >= 16.01 and self.vblétendreinbledur.value() <= 17:
                    self.rblétendreinbledur.setValue(0.425)
                elif self.vblétendreinbledur.value() >= 17.01 and self.vblétendreinbledur.value() <= 18:
                    self.rblétendreinbledur.setValue(0.550)
                elif self.vblétendreinbledur.value() >= 18.01 and self.vblétendreinbledur.value() <= 19:
                    self.rblétendreinbledur.setValue(0.675)
                elif self.vblétendreinbledur.value() >= 19.01 and self.vblétendreinbledur.value() <= 20:
                    self.rblétendreinbledur.setValue(0.825)
                elif self.vblétendreinbledur.value() >= 20.01 and self.vblétendreinbledur.value() <= 21:
                    self.rblétendreinbledur.setValue(0.975)
                elif self.vblétendreinbledur.value() >= 21.01 and self.vblétendreinbledur.value() <= 22:
                    self.rblétendreinbledur.setValue(1.150)
                elif self.vblétendreinbledur.value() >= 22.01 and self.vblétendreinbledur.value() <= 23:
                    self.rblétendreinbledur.setValue(1.325)
                elif self.vblétendreinbledur.value() >= 23.01 and self.vblétendreinbledur.value() <= 24:
                    self.rblétendreinbledur.setValue(1.525)
                elif self.vblétendreinbledur.value() >= 24.01 and self.vblétendreinbledur.value() <= 25:
                    self.rblétendreinbledur.setValue(1.70)
                elif self.vblétendreinbledur.value() >= 25.01 and self.vblétendreinbledur.value() <= 26:
                    self.rblétendreinbledur.setValue(1.90)
                elif self.vblétendreinbledur.value() >= 26.01 and self.vblétendreinbledur.value() <= 27:
                    self.rblétendreinbledur.setValue(2.10)
                elif self.vblétendreinbledur.value() >= 27.01 and self.vblétendreinbledur.value() <= 28:
                    self.rblétendreinbledur.setValue(2.30)
                elif self.vblétendreinbledur.value() >= 28.01 and self.vblétendreinbledur.value() <= 29:
                    self.rblétendreinbledur.setValue(2.50)
                elif self.vblétendreinbledur.value() >= 29.01 and self.vblétendreinbledur.value() <= 30:
                    self.rblétendreinbledur.setValue(2.75)
                elif self.vblétendreinbledur.value() >= 30.01 and self.vblétendreinbledur.value() <= 31:
                    self.rblétendreinbledur.setValue(3.00)
                elif self.vblétendreinbledur.value() >= 31.01 and self.vblétendreinbledur.value() <= 32:
                    self.rblétendreinbledur.setValue(3.25)
                elif self.vblétendreinbledur.value() >= 32.01 and self.vblétendreinbledur.value() <= 33:
                    self.rblétendreinbledur.setValue(3.50)
                elif self.vblétendreinbledur.value() >= 33.01 and self.vblétendreinbledur.value() <= 34:
                    self.rblétendreinbledur.setValue(3.75)
                elif self.vblétendreinbledur.value() >= 34.01 and self.vblétendreinbledur.value() <= 35:
                    self.rblétendreinbledur.setValue(4.00)
                else:
                    self.rblétendreinbledur.setValue(0)
                break

        def allcallculbd(self):
            self.psbd_calcull()
            self.ref_graincassébd()
            self.imp1erbd()
            self.imp2embd()
            self.plusbd()
            self.ref_grainbouté_bd()
            self.boni_bltdbd()

        def clear_allbd(self):
            self.vpsbd.clear()
            self.bpsbd.clear()
            self.vhumiditebd.clear()
            self.bhumiditebd.clear()
            self.vtotalprembd.clear()
            self.btotalprembd.clear()
            self.vgraincassébd.clear()
            self.rgraincassébd.clear()
            self.vtotaldembd.clear()
            self.btotalprembd.clear()
            self.rtotaldembd.clear()
            self.btotaldembd.clear()
            self.vgrainetrangébd.clear()
            self.vgrainmouchtébd.clear()
            self.vgrainmaigrebd.clear()
            self.vgrainechaudebd.clear()
            self.vdébrisvébd.clear()
            self.vmatiéreinrtbd.clear()
            self.btotalcomplet.clear()
            self.rtotalcomplet.clear()
            self.bindicenotin.clear()
            self.rindicenotin.clear()
            self.rblétendreinbledur.clear()
            self.rtotalprembd.clear()
            self.vdébrisvébd.clear()
            self.vmatiéreinrtbd.clear()
            self.vgrainsanvaleurbd.clear()
            self.vgrainchaufébd.clear()
            self.vgrainboutébd.clear()
            self.rgrainboutébd.clear()
            self.bblétendreinbledur.clear()
            self.vblétendreinbledur.clear()
            self.vtotalcomplet.clear()
            self.vindicenotin.clear()
            self.vgrainpiquébd.clear()
            self.vgrainpunaisébd.clear()
            self.vgraincarrébd.clear()
            self.vergotbd.clear()
            self.vergotbd.setValue(0)
            self.vgraigermébd.clear()
            self.vgraigermébd.setValue(0)
            self.vgrainnuisiblebd.clear()
            self.vgrainnuisiblebd.setValue(0)
            self.vgraincarrébd.setValue(0)
            self.vgrainpunaisébd.setValue(0)
            self.vgrainpiquébd.setValue(0)
            self.vindicenotin.setValue(0)
            self.vtotalcomplet.setValue(0)
            self.vpsbd.setValue(0)
            self.bpsbd.setValue(0)
            self.rpsbd.setValue(0)
            self.vhumiditebd.setValue(0)
            self.bhumiditebd.setValue(0)
            self.vtotalprembd.setValue(0)
            self.btotalprembd.setValue(0)
            self.vgraincassébd.setValue(0)
            self.rgraincassébd.setValue(0)
            self.vtotaldembd.setValue(0)
            self.btotalprembd.setValue(0)
            self.rtotaldembd.setValue(0)
            self.btotaldembd.setValue(0)
            self.vgrainetrangébd.setValue(0)
            self.vgrainmouchtébd.setValue(0)
            self.vgrainmaigrebd.setValue(0)
            self.vgrainechaudebd.setValue(0)
            self.vdébrisvébd.setValue(0)
            self.vmatiéreinrtbd.setValue(0)
            self.btotalcomplet.setValue(0)
            self.rtotalcomplet.setValue(0)
            self.bindicenotin.setValue(0)
            self.rindicenotin.setValue(0)
            self.rblétendreinbledur.setValue(0)
            self.rtotalprembd.setValue(0)
            self.vdébrisvébd.setValue(0)
            self.vmatiéreinrtbd.setValue(0)
            self.vgrainsanvaleurbd.setValue(0)
            self.vgrainchaufébd.setValue(0)
            self.vtotalprembd.setValue(0)
            self.btotalprembd.setValue(0)
            self.rtotalprembd.setValue(0)
            self.vgrainboutébd.setValue(0)
            self.rgrainboutébd.setValue(0)
            self.bblétendreinbledur.setValue(0)
            self.vblétendreinbledur.setValue(0)

        def plusbd(self):
            a = self.bpsbd.value()
            b = self.bhumiditebd.value()
            c = self.btotalprembd.value()
            a1 = self.bblétendreinbledur.value()
            d = self.rgraincassébd.value()
            e = self.vgraincassébd.value()
            f = self.vgrainmaigrebd.value()
            g = self.vgrainechaudebd.value()
            h = self.vgraigermébd.value()
            i = self.vgrainpunaisébd.value()
            j = self.vgrainpiquébd.value()
            k = self.vgrainboutébd.value()
            v = self.rgrainboutébd.value()
            m = self.vgrainmouchtébd.value()
            n = self.vgrainetrangébd.value()
            o = self.rtotalprembd.value()
            oa = self.rtotaldembd.value()
            t = self.rpsbd.value()
            q = self.vindicenotin.value()
            r = self.vblétendreinbledur.value()
            sa = self.btotaldembd.value()
            # total 2eme cat blé dur
            vtotaldemmbd = f + g + h + i + j + m + n
            # total 1er ble dur
            p = self.vdébrisvébd.value()
            qa = self.vmatiéreinrtbd.value()
            ra = self.vgrainchaufébd.value()
            s = self.vgrainsanvaleurbd.value()
            aa = self.vgrainnuisiblebd.value()
            ab = self.vgraincarrébd.value()
            if self.vgraincassébd.value() >= 4.01 and self.vgrainboutébd.value() >= 4.01:
                self.vtotaldembd.setValue(vtotaldemmbd)
                self.vtotalcomplet.setValue(vtotaldemmbd + q + r)
            elif self.vgraincassébd.value() >= 4.01 and self.vgrainboutébd.value() < 4.01:
                self.vtotaldembd.setValue(vtotaldemmbd + k)
                self.vtotalcomplet.setValue(vtotaldemmbd + q + r + k)
            elif self.vgrainboutébd.value() >= 4.01 and self.vgraincassébd.value() < 4.01:
                self.vtotaldembd.setValue(vtotaldemmbd + e)
                self.vtotalcomplet.setValue(vtotaldemmbd + q + r + e)
            else:
                self.vtotaldembd.setValue(vtotaldemmbd + e + k)
                self.vtotalcomplet.setValue(vtotaldemmbd + q + r + e + k)

            self.vtotalprembd.setValue(p + qa + ra + s + aa + ab)
            self.rtotalbonibd.setValue(t + o + d + v + oa)
            self.btotalbonibd.setValue(a + b + c + a1 + sa)

        def docx_bdsave(self):
            self.docbd = DocxTemplate("Docxfiles/_bulletin moulin/_bulletin moulin_Blé DUR/tempfile_blédur.docx")
            decadbd = self.decadecombobd.currentText()
            bbpsbd = self.bpsbd.text()
            ebpsbd = self.vpsbd.text()
            moulnbd = self.moulincombobd.currentText()
            pntclbd = self.pointcollectecombobd.currentText()
            espsbd = self.éspécecombobd.currentText()
            qnttbd = self.quantitetxtbd.text()
            agrcmbd = self.agréeeurcombobd.currentText()
            tnvvbd = self.vhumiditebd.text()
            ttv1bd = self.vtotalprembd.text()
            ttlb1bd = self.btotalprembd.text()
            ttlb1bdr = self.rtotalprembd.text()
            grcassévbd = self.vgraincassébd.text()
            grcassébbd = self.bgraincassébd.text()
            grcassébbdr = self.rgraincassébd.text()
            tnbbd = self.bhumiditebd.text()
            grmchbd = self.vgrainmouchtébd.value()
            gretrngbd = self.vgrainetrangébd.value()
            total2vbd = self.vtotaldembd.text()
            total2rbd = self.rtotaldembd.text()
            total2bbd = self.btotaldembd.text()
            grnmgrbd = self.vgrainmaigrebd.value()
            grechdbd = self.vgrainechaudebd.value()
            grgrmbd = self.vgraigermébd.value()
            grpnsbd = self.vgrainpunaisébd.value()
            grpqbd = self.vgrainpiquébd.value()
            grbtbd = self.vgrainboutébd.value()
            dattebd = self.dateeditebd.text()
            refactionps = self.rpsbd.text()
            debrivébd = self.vdébrisvébd.value()
            matinrtdb = self.vmatiéreinrtbd.value()
            grainmaigrebd = self.vgrainmouchtébd.value()
            grainboute = self.vgrainboutébd.value()
            garinssanvaleur = self.vgrainsanvaleurbd.value()
            grainchaufébd = self.vgrainchaufébd.value()
            indice = self.vindicenotin.value()
            indicer = self.rindicenotin.text()
            indiceb = self.bindicenotin.text()
            grainboutébdr = self.rgrainboutébd.text()
            grainboutébdb = self.bgrainboutébd.text()
            ttcomplet = self.vtotalcomplet.text()
            ttcompletr = self.rtotalcomplet.text()
            ttcompletb = self.btotalcomplet.text()
            bletendredbd = self.vblétendreinbledur.value()
            bletendredbdr = self.rblétendreinbledur.text()
            bletendredbdb = self.bblétendreinbledur.text()
            graincarre = self.vgraincarrébd.value()
            ergotbd = self.vergotbd.value()
            grainnuisiblebd = self.vgrainnuisiblebd.value()
            totalbonification = self.btotalbonibd.text()
            totalrefaction = self.rtotalbonibd.text()

            self.docbd.render(
                {"ttcr": totalrefaction, "ttcb": totalbonification, "gnsv": grainnuisiblebd, "erg": ergotbd,
                 "grcr": graincarre, "grbtf": grbtbd, "grpq": grpqbd, "grpn": grpnsbd, "grg": grgrmbd, "gehv": grechdbd,
                 "gmv": grnmgrbd,
                 "psb": bbpsbd, "tt1v": ttv1bd, "tneb": tnbbd, "tnev": tnvvbd, "psv": ebpsbd, "gr": agrcmbd,
                 "esp": espsbd, "mmmmmmmmmmmmmmmmm": moulnbd, "pntc": pntclbd, "qtt": qnttbd, "dt": dattebd,
                 "num": decadbd, "tt1b": ttlb1bd, "gcv": grcassévbd, "gcb": grcassébbd, "tt2v": total2vbd,
                 "tt2r": total2rbd, "tt2b": total2bbd, "getv": gretrngbd, "gmv": grmchbd, "gehv": grechdbd,
                 "gnsv": grainnuisiblebd,
                 "psr": refactionps,
                 "dbv": debrivébd,
                 "mtiv": matinrtdb,
                 "grch": grainchaufébd,
                 "grsn": garinssanvaleur,
                 "gmx": grainmaigrebd,
                 "grbt": grainboute,
                 "indv": indice,
                 "btdv": bletendredbd,
                 "ttcv": ttcomplet,
                 "ttcc": ttcompletr,
                 "ttcn": ttcompletb,
                 "btdr": bletendredbdr,
                 "btdb": bletendredbdb,
                 "indr": indicer,
                 "indb": indiceb,
                 "grbtr": grainboutébdr,
                 "grbtb": grainboutébdb,
                 "gcr": grcassébbdr,
                 "tt1r": ttlb1bdr})
            self.docbd_name = moulnbd + "-" + datetime.datetime.now().strftime("%m-%y") + "-" + decadbd + ".docx"
            # self.docbd.save("_bulletin moulin/_bulletin moulin_Blé DUR/" + self.docbd_name)
            path, _ = QFileDialog.getSaveFileName(None, "Enregistrer la fiche", self.docbd_name,
                                                  "Fichiers DOCX (*.docx)")
            if path:
                self.docbd.save(path)
                msgbox = QtWidgets.QMessageBox()
                msgbox.setWindowTitle('confirmation')
                msgbox.setText('Le fichier a été enregistré avec succès.')
                msgbox.exec()

        def printerbd(self):
            self.docbd = DocxTemplate("Docxfiles/_bulletin moulin/_bulletin moulin_Blé DUR/tempfile_blédur.docx")
            decadbd = self.decadecombobd.currentText()
            bbpsbd = self.bpsbd.text()
            ebpsbd = self.vpsbd.text()
            moulnbd = self.moulincombobd.currentText()
            pntclbd = self.pointcollectecombobd.currentText()
            espsbd = self.éspécecombobd.currentText()
            qnttbd = self.quantitetxtbd.text()
            agrcmbd = self.agréeeurcombobd.currentText()
            tnvvbd = self.vhumiditebd.text()
            ttv1bd = self.vtotalprembd.text()
            ttlb1bd = self.btotalprembd.text()
            ttlb1bdr = self.rtotalprembd.text()
            grcassévbd = self.vgraincassébd.text()
            grcassébbd = self.bgraincassébd.text()
            grcassébbdr = self.rgraincassébd.text()
            tnbbd = self.bhumiditebd.text()
            grmchbd = self.vgrainmouchtébd.value()
            gretrngbd = self.vgrainetrangébd.value()
            total2vbd = self.vtotaldembd.text()
            total2rbd = self.rtotaldembd.text()
            total2bbd = self.btotaldembd.text()
            grnmgrbd = self.vgrainmaigrebd.value()
            grechdbd = self.vgrainechaudebd.value()
            grgrmbd = self.vgraigermébd.value()
            grpnsbd = self.vgrainpunaisébd.value()
            grpqbd = self.vgrainpiquébd.value()
            grbtbd = self.vgrainboutébd.value()
            dattebd = self.dateeditebd.text()
            refactionps = self.rpsbd.text()
            debrivébd = self.vdébrisvébd.value()
            matinrtdb = self.vmatiéreinrtbd.value()
            grainmaigrebd = self.vgrainmouchtébd.value()
            grainboute = self.vgrainboutébd.value()
            garinssanvaleur = self.vgrainsanvaleurbd.value()
            grainchaufébd = self.vgrainchaufébd.value()
            indice = self.vindicenotin.value()
            indicer = self.rindicenotin.text()
            indiceb = self.bindicenotin.text()
            grainboutébdr = self.rgrainboutébd.text()
            grainboutébdb = self.bgrainboutébd.text()
            ttcomplet = self.vtotalcomplet.text()
            ttcompletr = self.rtotalcomplet.text()
            ttcompletb = self.btotalcomplet.text()
            bletendredbd = self.vblétendreinbledur.value()
            bletendredbdr = self.rblétendreinbledur.text()
            bletendredbdb = self.bblétendreinbledur.text()
            graincarre = self.vgraincarrébd.value()
            ergotbd = self.vergotbd.value()
            grainnuisiblebd = self.vgrainnuisiblebd.value()
            totalbonification = self.btotalbonibd.text()
            totalrefaction = self.rtotalbonibd.text()

            self.docbd.render(
                {"ttcr": totalrefaction, "ttcb": totalbonification, "gnsv": grainnuisiblebd, "erg": ergotbd,
                 "grcr": graincarre, "grbtf": grbtbd, "grpq": grpqbd, "grpn": grpnsbd, "grg": grgrmbd, "gehv": grechdbd,
                 "gmv": grnmgrbd,
                 "psb": bbpsbd, "tt1v": ttv1bd, "tneb": tnbbd, "tnev": tnvvbd, "psv": ebpsbd, "gr": agrcmbd,
                 "esp": espsbd, "mmmmmmmmmmmmmmmmm": moulnbd, "pntc": pntclbd, "qtt": qnttbd, "dt": dattebd,
                 "num": decadbd, "tt1b": ttlb1bd, "gcv": grcassévbd, "gcb": grcassébbd, "tt2v": total2vbd,
                 "tt2r": total2rbd, "tt2b": total2bbd, "getv": gretrngbd, "gmv": grmchbd, "gehv": grechdbd,
                 "gnsv": grainnuisiblebd,
                 "psr": refactionps,
                 "dbv": debrivébd,
                 "mtiv": matinrtdb,
                 "grch": grainchaufébd,
                 "grsn": garinssanvaleur,
                 "gmx": grainmaigrebd,
                 "grbt": grainboute,
                 "indv": indice,
                 "btdv": bletendredbd,
                 "ttcv": ttcomplet,
                 "ttcc": ttcompletr,
                 "ttcn": ttcompletb,
                 "btdr": bletendredbdr,
                 "btdb": bletendredbdb,
                 "indr": indicer,
                 "indb": indiceb,
                 "grbtr": grainboutébdr,
                 "grbtb": grainboutébdb,
                 "gcr": grcassébbdr,
                 "tt1r": ttlb1bdr})
            doc_names = tempfile.NamedTemporaryFile(suffix=".docx", delete=False).name
            doc_pdf = tempfile.NamedTemporaryFile(suffix=".pdf", delete=False).name
            self.docbd.save(doc_names)
            try:
                if doc_names:
                    a = self.progress_bar()
                    sys.stderr = open("consoleoutput.log", "w")
                    convert(doc_names, doc_pdf)
                    # Open the resulting .pdf file using the default associated application
                    os.startfile(doc_pdf, 'open')
                    #app_path = 'C:\\Program Files\\Okular\\bin\\okular.exe'
                    #subprocess.Popen([app_path, doc_pdf])
            except Exception as e:
                print(e)

        def progress_bar(self):
            self.widgetprogress = QtWidgets.QDialog()
            self.widgetprogress.setStyleSheet(""" QWidget
                                {
                                    color: #000000;
                                    background-color: #ffffff;
                                    border-width: 1px;
                                    border-color: #1e1e1e;
                                    border-style: solid;
                                    border-radius: 6;
                                    padding: 0px;
                                    font-size: 18px;
                                    padding-left: 1px;
                                    padding-right: 1px
                                }
                                QWidget:item:hover
                                {
                                    background-color: #3daee9;
                                    color: #eff0f1;
                                }
                                QWidget:item:selected
                                {
                                    background-color: #3daee9;
                                }
                                QWidget:disabled
                                {
                                    color: #454545;
                                    background-color: #31363b;
                                }
                                QPushButton
                                {
                                    color: #000000;
                                    background-color:#ade3e7;
                                    border-width: 1px;
                                    border-color: #1e1e1e;
                                    border-style: solid;
                                    border-radius: 6;
                                    padding: 3px;
                                    font-size: 12px;
                                    padding-left: 5px;
                                    padding-right: 5px;
                                    min-width: 40px
                                }
                                QPushButton:disabled
                                {
                                    background-color: #31363b;
                                    border-width: 1px;
                                    border-color: #454545;
                                    border-style: solid;
                                    padding-top: 5px;
                                    padding-bottom: 5px;
                                    padding-left: 10px;
                                    padding-right: 10px;
                                    border-radius: 2px;
                                    color: #454545;
                                }

                                QPushButton:pressed
                                {
                                    background-color: #3daee9;
                                    padding-top: -15px;
                                    padding-bottom: -17px;
                                }
                                QPushButton:hover
                                {
                                    border: 1px solid #ff8c00;
                                    color: #000000;
                                }
                                 QLabel
                                {
                                    font-size: 18px;
                                    border: 0px solid orange;
                                }

                            """)
            self.widgetprogress.setWindowTitle("جاري تحميل الملف يرجى الانتظار ")
            self.widgetprogress.setGeometry(550, 450, 250, 20)
            self.progressBar = QtWidgets.QProgressBar(self.widgetprogress)
            self.progressBar.setGeometry(10, 10, 200, 10)
            self.progressBar.setMinimum(0)
            self.progressBar.setMaximum(100)
            self.progressBar.setStyleSheet("""QProgressBar
        {
        border: solid grey;
        border-radius: 15px;
        color: black;
        }
        QProgressBar::chunk 
        {
        background-color: #05B8CC;
        border-radius :15px;
        }      """)
            self.progressBar.setAlignment(Qt.AlignmentFlag.AlignCenter)
            self.vbox = QVBoxLayout(self.widgetprogress)
            self.vbox.addWidget(self.progressBar)
            self.timer = QtCore.QTimer()
            self.timer.timeout.connect(self.update_progress)
            self.timer.start(5)  # Update progress every
            self.widgetprogress.show()

        def update_progress(self):
            # Simulate file download progress
            current_value = self.progressBar.value()
            if current_value < 100:
                new_value = current_value + 10
                self.progressBar.setValue(new_value)
                if current_value == 99:
                    self.timer.stop()
                    self.progressBar.close()
                    self.widgetprogress.close()

        #################blé tendre def
        def calcul_ps(self):
            while self.vps.value() >= 69:
                self.rps.setValue(2.60)
                if self.vps.value() <= 69.24:
                    self.rps.setValue(2.60)
                elif self.vps.value() >= 69.25 and self.vps.value() <= 69.49:
                    self.rps.setValue(2.50)
                elif self.vps.value() >= 69.50 and self.vps.value() <= 69.74:
                    self.rps.setValue(2.40)
                elif self.vps.value() >= 69.75 and self.vps.value() <= 69.99:
                    self.rps.setValue(2.30)
                elif self.vps.value() >= 70 and self.vps.value() <= 70.24:
                    self.rps.setValue(2.20)
                elif self.vps.value() >= 70.25 and self.vps.value() <= 70.49:
                    self.rps.setValue(2.10)
                elif self.vps.value() >= 70.50 and self.vps.value() <= 70.74:
                    self.rps.setValue(2.00)
                elif self.vps.value() >= 70.75 and self.vps.value() <= 70.99:
                    self.rps.setValue(1.90)
                elif self.vps.value() >= 71 and self.vps.value() <= 71.24:
                    self.rps.setValue(1.80)
                elif self.vps.value() >= 71.25 and self.vps.value() <= 71.49:
                    self.rps.setValue(1.70)
                elif self.vps.value() >= 71.50 and self.vps.value() <= 71.74:
                    self.rps.setValue(1.60)
                elif self.vps.value() >= 71.75 and self.vps.value() <= 71.99:
                    self.rps.setValue(1.50)
                elif self.vps.value() >= 72 and self.vps.value() <= 72.24:
                    self.rps.setValue(1.40)
                elif self.vps.value() >= 72.25 and self.vps.value() <= 72.49:
                    self.rps.setValue(1.30)
                elif self.vps.value() >= 72.50 and self.vps.value() <= 72.74:
                    self.rps.setValue(1.20)
                elif self.vps.value() >= 72.75 and self.vps.value() <= 72.99:
                    self.rps.setValue(1.10)
                elif self.vps.value() >= 73 and self.vps.value() <= 73.24:
                    self.rps.setValue(1.00)
                elif self.vps.value() >= 73.25 and self.vps.value() <= 73.49:
                    self.rps.setValue(0.90)
                elif self.vps.value() >= 73.50 and self.vps.value() <= 73.74:
                    self.rps.setValue(0.80)
                elif self.vps.value() >= 73.75 and self.vps.value() <= 73.99:
                    self.rps.setValue(0.70)
                elif self.vps.value() >= 74 and self.vps.value() <= 74.24:
                    self.rps.setValue(0.60)
                elif self.vps.value() >= 74.25 and self.vps.value() <= 74.49:
                    self.rps.setValue(0.50)
                elif self.vps.value() >= 74.50 and self.vps.value() <= 74.74:
                    self.rps.setValue(0.40)
                elif self.vps.value() >= 74.75 and self.vps.value() <= 74.99:
                    self.rps.setValue(0.30)
                elif self.vps.value() >= 75 and self.vps.value() <= 75.24:
                    self.rps.setValue(0.20)
                elif self.vps.value() >= 75.25 and self.vps.value() <= 75.49:
                    self.rps.setValue(0.10)
                else:
                    self.rps.setValue(0)
                if self.vps.value() >= 75.51 and self.vps.value() <= 75.75:
                    self.bps.setValue(0.10)
                elif self.vps.value() >= 75.76 and self.vps.value() <= 76:
                    self.bps.setValue(0.20)
                elif self.vps.value() >= 76.01 and self.vps.value() <= 76.25:
                    self.bps.setValue(0.30)
                elif self.vps.value() >= 76.26 and self.vps.value() <= 76.50:
                    self.bps.setValue(0.40)
                elif self.vps.value() >= 76.51 and self.vps.value() <= 76.75:
                    self.bps.setValue(0.50)
                elif self.vps.value() >= 77.76 and self.vps.value() <= 77:
                    self.bps.setValue(0.60)
                elif self.vps.value() >= 77.01 and self.vps.value() <= 77.25:
                    self.bps.setValue(0.70)
                elif self.vps.value() >= 77.26 and self.vps.value() <= 77.50:
                    self.bps.setValue(0.80)
                elif self.vps.value() >= 77.51 and self.vps.value() <= 77.75:
                    self.bps.setValue(0.90)
                elif self.vps.value() >= 77.76 and self.vps.value() <= 78:
                    self.bps.setValue(1.00)
                elif self.vps.value() >= 78.01 and self.vps.value() <= 78.25:
                    self.bps.setValue(1.05)
                elif self.vps.value() >= 78.26 and self.vps.value() <= 78.50:
                    self.bps.setValue(1.10)
                elif self.vps.value() >= 78.51 and self.vps.value() <= 78.75:
                    self.bps.setValue(1.15)
                elif self.vps.value() >= 78.76 and self.vps.value() <= 79:
                    self.bps.setValue(1.20)
                elif self.vps.value() >= 79.01 and self.vps.value() <= 79.25:
                    self.bps.setValue(1.25)
                elif self.vps.value() >= 79.26 and self.vps.value() <= 79.50:
                    self.bps.setValue(1.30)
                elif self.vps.value() >= 79.51 and self.vps.value() <= 79.75:
                    self.bps.setValue(1.35)
                elif self.vps.value() >= 79.76 and self.vps.value() <= 80:
                    self.bps.setValue(1.40)
                elif self.vps.value() >= 80.01 and self.vps.value() <= 80.25:
                    self.bps.setValue(1.42)
                elif self.vps.value() >= 80.26 and self.vps.value() <= 80.50:
                    self.bps.setValue(1.44)
                elif self.vps.value() >= 80.51 and self.vps.value() <= 80.75:
                    self.bps.setValue(1.46)
                elif self.vps.value() >= 80.76 and self.vps.value() <= 81:
                    self.bps.setValue(1.48)
                else:
                    self.bps.setValue(0)
                break

        def humidite_calcul(self):
            while self.vhumidite.value() >= 6.50:
                self.bhumidite.setValue(2.80)
                if self.vhumidite.value() <= 6.99:
                    self.bhumidite.setValue(2.80)
                elif self.vhumidite.value() >= 7 and self.vhumidite.value() <= 7.49:
                    self.bhumidite.setValue(2.60)
                elif self.vhumidite.value() >= 7.50 and self.vhumidite.value() <= 7.99:
                    self.bhumidite.setValue(2.40)
                elif self.vhumidite.value() >= 8.00 and self.vhumidite.value() <= 8.49:
                    self.bhumidite.setValue(2.20)
                elif self.vhumidite.value() >= 8.50 and self.vhumidite.value() <= 8.99:
                    self.bhumidite.setValue(2.00)
                elif self.vhumidite.value() >= 9 and self.vhumidite.value() <= 9.49:
                    self.bhumidite.setValue(1.80)
                elif self.vhumidite.value() >= 9.50 and self.vhumidite.value() <= 9.99:
                    self.bhumidite.setValue(1.60)
                elif self.vhumidite.value() >= 10 and self.vhumidite.value() <= 10.49:
                    self.bhumidite.setValue(1.40)
                elif self.vhumidite.value() >= 10.50 and self.vhumidite.value() <= 10.99:
                    self.bhumidite.setValue(1.20)
                elif self.vhumidite.value() >= 11 and self.vhumidite.value() <= 11.49:
                    self.bhumidite.setValue(1.00)
                elif self.vhumidite.value() >= 11.50 and self.vhumidite.value() <= 11.99:
                    self.bhumidite.setValue(0.80)
                elif self.vhumidite.value() >= 12 and self.vhumidite.value() <= 12.49:
                    self.bhumidite.setValue(0.60)
                elif self.vhumidite.value() >= 12.50 and self.vhumidite.value() <= 12.99:
                    self.bhumidite.setValue(0.40)
                elif self.vhumidite.value() >= 13 and self.vhumidite.value() <= 13.49:
                    self.bhumidite.setValue(0.20)
                else:
                    self.bhumidite.setValue(0.00)
                if self.vhumidite.value() >= 15.01 and self.vhumidite.value() <= 15.50:
                    self.rhumidite.setValue(0.20)
                elif self.vhumidite.value() >= 15.51 and self.vhumidite.value() <= 16:
                    self.rhumidite.setValue(0.40)
                elif self.vhumidite.value() >= 16.01 and self.vhumidite.value() <= 16.50:
                    self.rhumidite.setValue(0.60)
                elif self.vhumidite.value() >= 16.51 and self.vhumidite.value() <= 17:
                    self.rhumidite.setValue(0.80)
                elif self.vhumidite.value() >= 17.01 and self.vhumidite.value() <= 17.50:
                    self.rhumidite.setValue(1.00)
                elif self.vhumidite.value() >= 17.51 and self.vhumidite.value() <= 18:
                    self.rhumidite.setValue(1.20)
                else:
                    self.rhumidite.setValue(0)

                break

        def tota_er(self):
            while self.vtotalprem.value() >= 0.25:
                self.btotalprem.setValue(0.48)
                if self.vtotalprem.value() < 0.26:
                    self.btotalprem.setValue(0.48)
                elif self.vtotalprem.value() >= 0.26 and self.vtotalprem.value() <= 0.50:
                    self.btotalprem.setValue(0.36)
                elif self.vtotalprem.value() >= 0.51 and self.vtotalprem.value() <= 0.75:
                    self.btotalprem.setValue(0.24)
                elif self.vtotalprem.value() >= 0.76 and self.vtotalprem.value() <= 1.00:
                    self.btotalprem.setValue(0.12)
                else:
                    self.btotalprem.setValue(0)
                if self.vtotalprem.value() >= 1.01 and self.vtotalprem.value() <= 1.25:
                    self.rtotalprem.setValue(0.12)
                elif self.vtotalprem.value() >= 1.26 and self.vtotalprem.value() <= 1.50:
                    self.rtotalprem.setValue(0.24)
                elif self.vtotalprem.value() >= 1.51 and self.vtotalprem.value() <= 1.75:
                    self.rtotalprem.setValue(0.36)
                elif self.vtotalprem.value() >= 1.76 and self.vtotalprem.value() <= 2:
                    self.rtotalprem.setValue(0.48)
                elif self.vtotalprem.value() >= 2.01 and self.vtotalprem.value() <= 2.25:
                    self.rtotalprem.setValue(0.60)
                elif self.vtotalprem.value() >= 2.26 and self.vtotalprem.value() <= 2.50:
                    self.rtotalprem.setValue(0.72)
                elif self.vtotalprem.value() >= 2.51 and self.vtotalprem.value() <= 2.75:
                    self.rtotalprem.setValue(0.84)
                elif self.vtotalprem.value() >= 2.76 and self.vtotalprem.value() <= 3:
                    self.rtotalprem.setValue(0.96)
                elif self.vtotalprem.value() >= 3.01 and self.vtotalprem.value() <= 3.25:
                    self.rtotalprem.setValue(1.08)
                elif self.vtotalprem.value() >= 3.26 and self.vtotalprem.value() <= 3.50:
                    self.rtotalprem.setValue(1.20)
                elif self.vtotalprem.value() >= 3.51 and self.vtotalprem.value() <= 3.75:
                    self.rtotalprem.setValue(1.32)
                elif self.vtotalprem.value() >= 3.76 and self.vtotalprem.value() <= 4:
                    self.rtotalprem.setValue(1.44)
                elif self.vtotalprem.value() >= 4.01 and self.vtotalprem.value() <= 4.25:
                    self.rtotalprem.setValue(1.56)
                elif self.vtotalprem.value() >= 4.26 and self.vtotalprem.value() <= 4.50:
                    self.rtotalprem.setValue(1.68)
                elif self.vtotalprem.value() >= 4.51 and self.vtotalprem.value() <= 4.75:
                    self.rtotalprem.setValue(1.80)
                elif self.vtotalprem.value() >= 4.76 and self.vtotalprem.value() <= 5:
                    self.rtotalprem.setValue(1.92)
                else:
                    self.rtotalprem.setValue(0)
                break

        def grain_cassé(self):
            while self.vgraincassé.value() >= 2.01:
                self.rgraincassé.setValue(0.04)
                if self.vgraincassé.value() <= 2.25:
                    self.rgraincassé.setValue(0.04)
                elif self.vgraincassé.value() >= 2.26 and self.vgraincassé.value() <= 2.50:
                    self.rgraincassé.setValue(0.08)
                elif self.vgraincassé.value() >= 2.51 and self.vgraincassé.value() <= 2.75:
                    self.rgraincassé.setValue(0.12)
                elif self.vgraincassé.value() >= 2.76 and self.vgraincassé.value() <= 3.00:
                    self.rgraincassé.setValue(0.16)
                elif self.vgraincassé.value() >= 3.01 and self.vgraincassé.value() <= 3.25:
                    self.rgraincassé.setValue(0.20)
                elif self.vgraincassé.value() >= 3.26 and self.vgraincassé.value() <= 3.50:
                    self.rgraincassé.setValue(0.24)
                elif self.vgraincassé.value() >= 3.51 and self.vgraincassé.value() <= 3.75:
                    self.rgraincassé.setValue(0.28)
                elif self.vgraincassé.value() >= 3.76 and self.vgraincassé.value() <= 4.00:
                    self.rgraincassé.setValue(0.32)
                elif self.vgraincassé.value() >= 4.01 and self.vgraincassé.value() <= 4.25:
                    self.rgraincassé.setValue(0.36)
                elif self.vgraincassé.value() >= 4.26 and self.vgraincassé.value() <= 4.50:
                    self.rgraincassé.setValue(0.40)
                elif self.vgraincassé.value() >= 4.51 and self.vgraincassé.value() <= 4.75:
                    self.rgraincassé.setValue(0.44)
                elif self.vgraincassé.value() >= 4.76 and self.vgraincassé.value() <= 5.00:
                    self.rgraincassé.setValue(0.48)
                elif self.vgraincassé.value() >= 5.01 and self.vgraincassé.value() <= 5.25:
                    self.rgraincassé.setValue(0.54)
                elif self.vgraincassé.value() >= 5.26 and self.vgraincassé.value() <= 5.50:
                    self.rgraincassé.setValue(0.60)
                elif self.vgraincassé.value() >= 5.51 and self.vgraincassé.value() <= 5.75:
                    self.rgraincassé.setValue(0.66)
                elif self.vgraincassé.value() >= 5.76 and self.vgraincassé.value() <= 6.00:
                    self.rgraincassé.setValue(0.72)
                elif self.vgraincassé.value() >= 6.01 and self.vgraincassé.value() <= 6.25:
                    self.rgraincassé.setValue(0.78)
                elif self.vgraincassé.value() >= 6.26 and self.vgraincassé.value() <= 6.50:
                    self.rgraincassé.setValue(0.84)
                elif self.vgraincassé.value() >= 6.51 and self.vgraincassé.value() <= 6.75:
                    self.rgraincassé.setValue(0.90)
                elif self.vgraincassé.value() >= 6.76 and self.vgraincassé.value() <= 7.00:
                    self.rgraincassé.setValue(0.96)
                elif self.vgraincassé.value() >= 7.01 and self.vgraincassé.value() <= 7.25:
                    self.rgraincassé.setValue(1.02)
                elif self.vgraincassé.value() >= 7.26 and self.vgraincassé.value() <= 7.50:
                    self.rgraincassé.setValue(1.08)
                elif self.vgraincassé.value() >= 7.51 and self.vgraincassé.value() <= 7.75:
                    self.rgraincassé.setValue(1.14)
                elif self.vgraincassé.value() >= 7.76 and self.vgraincassé.value() <= 8.00:
                    self.rgraincassé.setValue(1.20)
                else:
                    self.rgraincassé.setValue(0)
                break

        def total_eme(self):
            while self.vtotaldem.value() >= 5.01:
                self.rtotaldem.setValue(0.05)
                if self.vtotaldem.value() <= 5.25:
                    self.rtotaldem.setValue(0.05)
                elif self.vtotaldem.value() >= 5.26 and self.vtotaldem.value() <= 5.50:
                    self.rtotaldem.setValue(0.10)
                elif self.vtotaldem.value() >= 5.51 and self.vtotaldem.value() <= 5.75:
                    self.rtotaldem.setValue(0.15)
                elif self.vtotaldem.value() >= 5.76 and self.vtotaldem.value() <= 6:
                    self.rtotaldem.setValue(0.20)
                elif self.vtotaldem.value() >= 6.01 and self.vtotaldem.value() <= 6.25:
                    self.rtotaldem.setValue(0.25)
                elif self.vtotaldem.value() >= 6.26 and self.vtotaldem.value() <= 6.50:
                    self.rtotaldem.setValue(0.30)
                elif self.vtotaldem.value() >= 6.51 and self.vtotaldem.value() <= 6.75:
                    self.rtotaldem.setValue(0.35)
                elif self.vtotaldem.value() >= 6.76 and self.vtotaldem.value() <= 7:
                    self.rtotaldem.setValue(0.40)
                elif self.vtotaldem.value() >= 7.01 and self.vtotaldem.value() <= 7.25:
                    self.rtotaldem.setValue(0.45)
                elif self.vtotaldem.value() >= 7.26 and self.vtotaldem.value() <= 7.50:
                    self.rtotaldem.setValue(0.50)
                elif self.vtotaldem.value() >= 7.51 and self.vtotaldem.value() <= 7.75:
                    self.rtotaldem.setValue(0.55)
                elif self.vtotaldem.value() >= 7.76 and self.vtotaldem.value() <= 8:
                    self.rtotaldem.setValue(0.60)
                elif self.vtotaldem.value() >= 8.01 and self.vtotaldem.value() <= 8.25:
                    self.rtotaldem.setValue(0.65)
                elif self.vtotaldem.value() >= 8.26 and self.vtotaldem.value() <= 8.50:
                    self.rtotaldem.setValue(0.70)
                elif self.vtotaldem.value() >= 8.51 and self.vtotaldem.value() <= 8.75:
                    self.rtotaldem.setValue(0.75)
                elif self.vtotaldem.value() >= 8.76 and self.vtotaldem.value() <= 9:
                    self.rtotaldem.setValue(0.80)
                elif self.vtotaldem.value() >= 9.01 and self.vtotaldem.value() <= 9.25:
                    self.rtotaldem.setValue(0.85)
                elif self.vtotaldem.value() >= 9.26 and self.vtotaldem.value() <= 9.50:
                    self.rtotaldem.setValue(0.90)
                elif self.vtotaldem.value() >= 9.51 and self.vtotaldem.value() <= 9.75:
                    self.rtotaldem.setValue(0.95)
                elif self.vtotaldem.value() >= 9.76 and self.vtotaldem.value() <= 10:
                    self.rtotaldem.setValue(1.00)
                elif self.vtotaldem.value() >= 10.01 and self.vtotaldem.value() <= 10.25:
                    self.rtotaldem.setValue(1.08)
                elif self.vtotaldem.value() >= 10.26 and self.vtotaldem.value() <= 10.50:
                    self.rtotaldem.setValue(1.16)
                elif self.vtotaldem.value() >= 10.51 and self.vtotaldem.value() <= 10.75:
                    self.rtotaldem.setValue(1.24)
                elif self.vtotaldem.value() >= 10.76 and self.vtotaldem.value() <= 11:
                    self.rtotaldem.setValue(1.32)
                elif self.vtotaldem.value() >= 11.01 and self.vtotaldem.value() <= 11.25:
                    self.rtotaldem.setValue(1.40)
                elif self.vtotaldem.value() >= 11.26 and self.vtotaldem.value() <= 11.50:
                    self.rtotaldem.setValue(1.48)
                elif self.vtotaldem.value() >= 11.51 and self.vtotaldem.value() <= 11.75:
                    self.rtotaldem.setValue(1.56)
                elif self.vtotaldem.value() >= 11.76 and self.vtotaldem.value() <= 12:
                    self.rtotaldem.setValue(1.64)
                elif self.vtotaldem.value() >= 12.01 and self.vtotaldem.value() <= 12.25:
                    self.rtotalprem.setValue(1.72)
                elif self.vtotaldem.value() >= 12.26 and self.vtotaldem.value() <= 12.50:
                    self.rtotaldem.setValue(1.80)
                elif self.vtotaldem.value() >= 12.51 and self.vtotaldem.value() <= 12.75:
                    self.rtotaldem.setValue(1.88)
                elif self.vtotaldem.value() >= 12.76 and self.vtotaldem.value() <= 13:
                    self.rtotaldem.setValue(1.96)
                elif self.vtotaldem.value() >= 13.01 and self.vtotaldem.value() <= 13.25:
                    self.rtotalprem.setValue(2.04)
                elif self.vtotaldem.value() >= 13.26 and self.vtotaldem.value() <= 13.50:
                    self.rtotaldem.setValue(2.12)
                elif self.vtotaldem.value() >= 13.51 and self.vtotaldem.value() <= 13.75:
                    self.rtotaldem.setValue(2.20)
                elif self.vtotaldem.value() >= 13.76 and self.vtotaldem.value() <= 14:
                    self.rtotaldem.setValue(2.28)
                elif self.vtotaldem.value() >= 14.01 and self.vtotaldem.value() <= 14.25:
                    self.rtotalprem.setValue(2.36)
                elif self.vtotaldem.value() >= 14.26 and self.vtotaldem.value() <= 14.50:
                    self.rtotaldem.setValue(2.44)
                elif self.vtotaldem.value() >= 14.51 and self.vtotaldem.value() <= 14.75:
                    self.rtotaldem.setValue(2.52)
                elif self.vtotaldem.value() >= 14.76 and self.vtotaldem.value() <= 15:
                    self.rtotaldem.setValue(2.60)
                else:
                    self.rtotaldem.value(0)
                break

        def plus(self):
            a = self.bps.value()
            b = self.bhumidite.value()
            c = self.btotalprem.value()
            c1 = self.btotaldem.value()
            d = self.rgraincassé.value()
            e = self.vgraincassé.value()
            f = self.vgrainmaigre.value()
            g = self.vgrainechaude.value()
            h = self.vgraigermé.value()
            i = self.vgrainpunaisé.value()
            j = self.vgrainpiqué.value()
            k = self.vgrainbouté.value()
            l = self.vgrainboutef.value()
            m = self.vgrainmouchté.value()
            n = self.vgrainetrangé.value()
            o = self.rtotaldem.value()
            t = self.rtotalprem.value()
            p = self.vdébrisvé.value()
            q = self.vmatiéreinrt.value()
            r = self.vgrainchaufé.value()
            s = self.vgrainsanvaleur.value()
            v = self.vgraincarré.value()
            aa = self.rps.value()
            ab = self.rhumidite.value()
            ac = self.vgrainnuisible.value()
            vtotaldemm = f + g + h + i + j + k + l + m + n
            totalbonifica = a + b + c + c1
            totalrefaction = o + t + d
            self.rtotalboni.setValue(totalrefaction)
            self.btotalboni.setValue(totalbonifica)
            self.vtotalprem.setValue(p + q + r + s + v + ac)
            self.vtotaldem.setValue(vtotaldemm)
            if self.vgraincassé.value() >= 2.01 and self.vgraincassé.value() <= 15:
                self.vtotaldem.setValue(vtotaldemm)
            else:
                self.vtotaldem.setValue(e + vtotaldemm)

            # vtotalpremm=p+q+r+s

        def all_calcul(self):
            self.tota_er()
            self.total_eme()
            self.calcul_ps()
            self.humidite_calcul()
            self.grain_cassé()
            self.plus()

        def clear_all(self):
            self.vgrainnuisible.clear()
            self.vgrainnuisible.setValue(0)
            self.vgrainchaufé.clear()
            self.vgrainchaufé.setValue(0)
            self.vps.clear()
            self.bps.clear()
            self.rps.clear()
            self.vgrainboutef.clear()
            self.vgrainbouté.clear()
            self.vhumidite.clear()
            self.bhumidite.clear()
            self.vgrainsanvaleur.clear()
            self.vgraincarré.clear()
            self.vtotalprem.clear()
            self.btotalprem.clear()
            self.vgraincassé.clear()
            self.rgraincassé.clear()
            self.vtotaldem.clear()
            self.btotalprem.clear()
            self.rtotaldem.clear()
            self.btotaldem.clear()
            self.vgrainpunaisé.clear()
            self.vgrainpunaisé.setValue(0)
            self.vgraigermé.clear()
            self.vgraigermé.setValue(0)
            self.vgrainetrangé.clear()
            self.vgrainmouchté.clear()
            self.vgrainmaigre.clear()
            self.vgrainechaude.clear()
            self.vgrainbouté.clear()
            self.vdébrisvé.clear()
            self.vmatiéreinrt.clear()
            self.vgrainpiqué.clear()
            self.vgrainpiqué.setValue(0)
            self.vps.setValue(0)
            self.bps.setValue(0)
            self.vhumidite.setValue(0)
            self.bhumidite.setValue(0)
            self.vgrainsanvaleur.setValue(0)
            self.vgraincarré.setValue(0)
            self.vtotalprem.setValue(0)
            self.btotalprem.setValue(0)
            self.vgraincassé.setValue(0)
            self.vgrainboutef.setValue(0)
            self.rgraincassé.setValue(0)
            self.vtotaldem.setValue(0)
            self.btotalprem.setValue(0)
            self.rtotaldem.setValue(0)
            self.btotaldem.setValue(0)
            self.vgrainetrangé.setValue(0)
            self.vgrainmouchté.setValue(0)
            self.vgrainmaigre.setValue(0)
            self.vgrainechaude.setValue(0)
            self.vdébrisvé.setValue(0)
            self.vmatiéreinrt.setValue(0)
            self.vgrainbouté.setValue(0)

        def docx_file(self):
            self.doc = DocxTemplate("Docxfiles/_bulletin moulin/_bulletin moulin_Blé TENDRE/invoice_template.docx")
            decad = self.decadecombo.currentText()
            bbps = self.bps.text()
            ebps = self.vps.text()
            ergot = self.vergot.value()
            nuisible = self.vgrainnuisible.value()
            mouln = self.moulincombo.currentText()
            pntcl = self.pointcollectecombo.currentText()
            esps = self.éspécecombo.currentText()
            qntt = self.quantitetxt.text()
            agrcm = self.agréeeurcombo.currentText()
            tnvv = self.vhumidite.text()
            debritvegetaux = self.vdébrisvé.value()
            matierinert = self.vmatiéreinrt.value()
            grainchaufe = self.vgrainchaufé.value()
            grainsanvaleur = self.vgrainsanvaleur.value()
            graincarie = self.vgraincarré.value()
            graingerme = self.vgraigermé.value()
            grainpunaisés = self.vgrainpunaisé.value()
            ttv1 = self.vtotalprem.text()
            ttlb1 = self.btotalprem.text()
            ttlr1 = self.rtotalprem.text()
            grcassév = self.vgraincassé.text()
            grcasséb = self.bgraincassé.text()
            rfgss = self.rgraincassé.text()
            tnb = self.bhumidite.text()
            grmgre = self.vgrainmaigre.text()
            grechd = self.vgrainechaude.text()
            grmch = self.vgrainmouchté.text()
            gretrng = self.vgrainetrangé.text()
            total2v = self.vtotaldem.text()
            total2r = self.rtotaldem.text()
            total2b = self.btotaldem.text()
            grnmgr = self.vgrainmaigre.value()
            grechd = self.vgrainechaude.value()
            grgrm = self.vgraigermé.value()
            grpns = self.vgrainpunaisé.value()
            grpq = self.vgrainpiqué.value()
            grbt = self.vgrainbouté.value()
            grnbtf = self.vgrainboutef.value()
            totalbonification = self.btotalboni.value()
            totalrefaction = self.rtotalboni.value()
            datte = self.dateedite.text()
            self.doc.render(
                {"tt1r": ttlr1, "tttb": totalbonification, "tttr": totalrefaction, "verg": ergot, "vnsb": nuisible,
                 "grpn": grainpunaisés, "grch": grainchaufe, "grsn": grainsanvaleur, "vgc": graincarie,
                 "grg": graingerme, "grbt": grnbtf, "grbtf": grbt, "dbv": debritvegetaux, "mtiv": matierinert,
                 "grpq": grpq, "grpn": grpns, "grg": grgrm, "gehv": grechd, "gmx": grnmgr, "psb": bbps, "tt1v": ttv1,
                 "tneb": tnb, "tnev": tnvv, "psv": ebps, "gr": agrcm, "esp": esps, "mmmmmmmmmmmmmmmmm": mouln,
                 "pntc": pntcl, "qtt": qntt, "dt": datte, "num": decad, "tt1b": ttlb1, "gcv": grcassév, "gcb": grcasséb,
                 "gcr": rfgss, "tt2v": total2v, "tt2r": total2r, "tt2b": total2b, "getv": gretrng, "gmv": grmch,
                 "gehv": grechd, "gmx": grmgre})
            doc_name = mouln + "-" + datetime.datetime.now().strftime("%m-%y") + "-" + decad + ".docx"
            # self.doc.save("_bulletin moulin/_bulletin moulin_Blé TENDRE/" + self.doc_name)
            path, _ = QFileDialog.getSaveFileName(None, "Enregistrer la fiche", doc_name, "Fichiers DOCX (*.docx)")
            if path:
                self.doc.save(path)
                msgbox = QtWidgets.QMessageBox()
                msgbox.setWindowTitle('confirmation')
                msgbox.setText('Le fichier a été enregistré avec succès.')
                msgbox.exec()

        def printer(self):
            self.doc = DocxTemplate("Docxfiles/_bulletin moulin/_bulletin moulin_Blé TENDRE/invoice_template.docx")
            decad = self.decadecombo.currentText()
            bbps = self.bps.text()
            ebps = self.vps.text()
            ergot = self.vergot.value()
            nuisible = self.vgrainnuisible.value()
            mouln = self.moulincombo.currentText()
            pntcl = self.pointcollectecombo.currentText()
            esps = self.éspécecombo.currentText()
            qntt = self.quantitetxt.text()
            agrcm = self.agréeeurcombo.currentText()
            tnvv = self.vhumidite.text()
            debritvegetaux = self.vdébrisvé.value()
            matierinert = self.vmatiéreinrt.value()
            grainchaufe = self.vgrainchaufé.value()
            grainsanvaleur = self.vgrainsanvaleur.value()
            graincarie = self.vgraincarré.value()
            graingerme = self.vgraigermé.value()
            grainpunaisés = self.vgrainpunaisé.value()
            ttv1 = self.vtotalprem.text()
            ttlb1 = self.btotalprem.text()
            ttlr1 = self.rtotalprem.text()
            grcassév = self.vgraincassé.text()
            grcasséb = self.bgraincassé.text()
            rfgss = self.rgraincassé.text()
            tnb = self.bhumidite.text()
            grmgre = self.vgrainmaigre.value()
            grechd = self.vgrainechaude.value()
            grmch = self.vgrainmouchté.value()
            gretrng = self.vgrainetrangé.text()
            total2v = self.vtotaldem.text()
            total2r = self.rtotaldem.text()
            total2b = self.btotaldem.text()
            grnmgr = self.vgrainmaigre.value()
            grechd = self.vgrainechaude.value()
            grgrm = self.vgraigermé.value()
            grpns = self.vgrainpunaisé.value()
            grpq = self.vgrainpiqué.value()
            grbt = self.vgrainbouté.value()
            grnbtf = self.vgrainboutef.value()
            totalbonification = self.btotalboni.value()
            totalrefaction = self.rtotalboni.value()
            datte = self.dateedite.text()
            self.doc.render(
                {"tt1r": ttlr1, "tttb": totalbonification, "tttr": totalrefaction, "verg": ergot, "vnsb": nuisible,
                 "grpn": grainpunaisés, "grch": grainchaufe, "grsn": grainsanvaleur, "vgc": graincarie,
                 "grg": graingerme,
                 "grbt": grnbtf, "grbtf": grbt, "dbv": debritvegetaux, "mtiv": matierinert, "grpq": grpq, "grpn": grpns,
                 "grg": grgrm, "gehv": grechd, "gmx": grnmgr, "psb": bbps, "tt1v": ttv1, "tneb": tnb, "tnev": tnvv,
                 "psv": ebps, "gr": agrcm, "esp": esps, "mmmmmmmmmmmmmmmmm": mouln, "pntc": pntcl, "qtt": qntt,
                 "dt": datte,
                 "num": decad, "tt1b": ttlb1, "gcv": grcassév, "gcb": grcasséb, "gcr": rfgss, "tt2v": total2v,
                 "tt2r": total2r, "tt2b": total2b, "getv": gretrng, "gmv": grmch, "gehv": grechd, "gmv": grmgre})
            self.doc_name = mouln + "-" + datetime.datetime.now().strftime("%m") + "-" + decad + ".docx"
            doc_names = tempfile.NamedTemporaryFile(suffix=".docx", delete=False).name
            doc_pdf = tempfile.NamedTemporaryFile(suffix=".pdf", delete=False).name
            self.doc.save(doc_names)
            try:
                if doc_names:
                    a = self.progress_bar()
                    sys.stderr = open("consoleoutput.log", "w")
                    convert(doc_names, doc_pdf)
                    # Open the resulting .pdf file using the default associated application
                    os.startfile(doc_pdf, 'open')
                    #app_path = 'C:\\Program Files\\Okular\\bin\\okular.exe'
                    #subprocess.Popen([app_path, doc_pdf])
            except Exception as e:
                print(e)

        def mouli_window(self, MainWindow):
            MainWindow.setObjectName("MainWindow")
            MainWindow.resize(1350, 700)
            MainWindow.setWindowIcon(QIcon("images/Picsart_23-03-13_18-53-05-983.ico"))
            MainWindow.setStyleSheet("""QToolTip
            {
                border: 1px solid #76797C;
                background-color: #fff8b0;
                color: white;
                padding: 5px;
                opacity: 200;
            }
            QWidget
            {
                color: #eff0f1;
                background-color: #ffffff;
                selection-background-color:#3daee9;
                selection-color: #eff0f1;
                background-clip: border;
                border-image: none;
                border: 0px transparent black;
                outline: 0;
            }

            QWidget:item:hover
            {
                background-color: #3daee9;
                color: #eff0f1;
            }

            QWidget:item:selected
            {
                background-color: #3daee9;
            }



            QWidget:disabled
            {
                color: #454545;
                background-color: #31363b;
            }

            QAbstractItemView
            {
                alternate-background-color: #31363b;
                color: #eff0f1;
                border: 1px solid 3A3939;
                border-radius: 2px;
            }

            QWidget:focus, QMenuBar:focus
            {
                border: 1px solid #3daee9;
            }
            QTabWidget::pane 
            {
             border: 2px solid red; 
             }

            QTabWidget:focus, QCheckBox:focus, QRadioButton:focus, QSlider:focus
            {
                border: none;
            }

            QLineEdit
            {
                background-color: #232629;
                padding: 1px;
                border-style: solid;
                border: 1px solid #76797C;
                border-radius: 2px;
                color: #eff0f1;
            }
            QDoubleSpinBox
            {
                background-color: #232629;
                padding: 1px;
                border-style: solid;
                border: 1px solid #76797C;
                border-radius: 2px;
                color:#eff0f1;

            }
            QDoubleSpinBox::drop-down
            {
                subcontrol-origin: padding;
                subcontrol-position: top right;
                width: 1px;

                border-left-width: 0px;
                border-left-color: #232629;
                border-left-style: solid;
                border-top-right-radius: 1px;
                border-bottom-right-radius: 1px;
            }



            QGroupBox {
                border:1px solid #76797C;
                border-radius: 2px;
                margin-top: 20px;
            }

            QGroupBox::title {
                subcontrol-origin: margin;
                subcontrol-position: top center;
                padding-left: 10px;
                padding-right: 10px;
                padding-top: 10px;
            }

            QAbstractScrollArea
            {
                border-radius: 2px;
                border: 1px solid #76797C;
                background-color: transparent;
            }

            QScrollBar:horizontal
            {
                height: 15px;
                margin: 3px 15px 3px 15px;
                border: 1px transparent #2A2929;
                border-radius: 4px;
                background-color: #2A2929;
            }

            QScrollBar::handle:horizontal
            {
                background-color: #605F5F;
                min-width: 5px;
                border-radius: 4px;
            }

            QScrollBar::add-line:horizontal
            {
                margin: 0px 3px 0px 3px;
                border-image: url(:/qss_icons/Dark_rc/right_arrow_disabled.png);
                width: 10px;
                height: 10px;
                subcontrol-position: right;
                subcontrol-origin: margin;
            }

            QScrollBar::sub-line:horizontal
            {
                margin: 0px 3px 0px 3px;
                border-image: url(:/qss_icons/Dark_rc/left_arrow_disabled.png);
                height: 10px;
                width: 10px;
                subcontrol-position: left;
                subcontrol-origin: margin;
            }

            QScrollBar::add-line:horizontal:hover,QScrollBar::add-line:horizontal:on
            {
                border-image: url(:/qss_icons/Dark_rc/right_arrow.png);
                height: 10px;
                width: 10px;
                subcontrol-position: right;
                subcontrol-origin: margin;
            }


            QScrollBar::sub-line:horizontal:hover, QScrollBar::sub-line:horizontal:on
            {
                border-image: url(:/qss_icons/Dark_rc/left_arrow.png);
                height: 10px;
                width: 10px;
                subcontrol-position: left;
                subcontrol-origin: margin;
            }

            QScrollBar::up-arrow:horizontal, QScrollBar::down-arrow:horizontal
            {
                background: none;
            }


            QScrollBar::add-page:horizontal, QScrollBar::sub-page:horizontal
            {
                background: none;
            }

            QScrollBar:vertical
            {
                background-color: #2A2929;
                width: 15px;
                margin: 15px 3px 15px 3px;
                border: 1px transparent #2A2929;
                border-radius: 4px;
            }

            QScrollBar::handle:vertical
            {
                background-color: #605F5F;
                min-height: 5px;
                border-radius: 4px;
            }

            QScrollBar::sub-line:vertical
            {
                margin: 3px 0px 3px 0px;
                border-image: url(:/qss_icons/Dark_rc/up_arrow_disabled.png);
                height: 10px;
                width: 10px;
                subcontrol-position: top;
                subcontrol-origin: margin;
            }

            QScrollBar::add-line:vertical
            {
                margin: 3px 0px 3px 0px;
                border-image: url(:/qss_icons/Dark_rc/down_arrow_disabled.png);
                height: 10px;
                width: 10px;
                subcontrol-position: bottom;
                subcontrol-origin: margin;
            }

            QScrollBar::sub-line:vertical:hover,QScrollBar::sub-line:vertical:on
            {

                border-image: url(:/qss_icons/Dark_rc/up_arrow.png);
                height: 10px;
                width: 10px;
                subcontrol-position: top;
                subcontrol-origin: margin;
            }


            QScrollBar::add-line:vertical:hover, QScrollBar::add-line:vertical:on
            {
                border-image: url(:/qss_icons/Dark_rc/down_arrow.png);
                height: 10px;
                width: 10px;
                subcontrol-position: bottom;
                subcontrol-origin: margin;
            }

            QScrollBar::up-arrow:vertical, QScrollBar::down-arrow:vertical
            {
                background: none;
            }


            QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical
            {
                background: none;
            }

            QTextEdit
            {
                background-color: #232629;
                color: #eff0f1;
                border: 1px solid #76797C;
            }

            QPlainTextEdit
            {
                background-color: #232629;;
                color: #eff0f1;
                border-radius: 2px;
                border: 1px solid #76797C;
            }

            QHeaderView::section
            {
                background-color: #76797C;
                color: #eff0f1;
                padding: 5px;
                border: 1px solid #76797C;
            }

            QSizeGrip {
                image: url(:/qss_icons/Dark_rc/sizegrip.png);
                width: 12px;
                height: 12px;
            }


            QMainWindow::separator
            {
                background-color: #31363b;
                color: white;
                padding-left: 4px;
                spacing: 2px;
                border: 1px dashed #76797C;
            }

            QMainWindow::separator:hover
            {

                background-color: #787876;
                color: white;
                padding-left: 4px;
                border: 1px solid #76797C;
                spacing: 2px;
            }


            QMenu::separator
            {
                height: 1px;
                background-color: #76797C;
                color: white;
                padding-left: 4px;
                margin-left: 10px;
                margin-right: 5px;
            }


            QFrame
            {
                border-radius: 2px;
                border: 1px solid #76797C;
            }

            QFrame[frameShape="0"]
            {
                border-radius: 2px;
                border: 1px transparent #76797C;
            }

            QStackedWidget
            {
                border: 1px transparent black;
            }


            QPushButton
            {
                color: #b1b1b1;
                background-color: QLinearGradient( x1: 0, y1: 0, x2: 0, y2: 1, stop: 0 #565656, stop: 0.1 #525252, stop: 0.5 #4e4e4e, stop: 0.9 #4a4a4a, stop: 1 #464646);
                border-width: 1px;
                border-color: #1e1e1e;
                border-style: solid;
                border-radius: 6;
                padding: 3px;
                font-size: 12px;
                padding-left: 5px;
                padding-right: 5px;
                min-width: 40px;

            }

            QPushButton:disabled
            {
                background-color: #31363b;
                border-width: 1px;
                border-color: #454545;
                border-style: solid;
                padding-top: 5px;
                padding-bottom: 5px;
                padding-left: 10px;
                padding-right: 10px;
                border-radius: 2px;
                color: #454545;
            }

            QPushButton:focus {
                background-color: #3daee9;
                color: white;
            }

            QPushButton:pressed
            {
                background-color: #3daee9;
                padding-top: -15px;
                padding-bottom: -17px;
            }

            

            QPushButton:checked{
                background-color: #76797C;
                border-color: #6A6969;
            }

            QComboBox {
    background-color: #ffffff;
    border: 1px solid #76797C;
    color:#000000;
    border-radius: 0.25em;
    padding: 0.10em 0.10em;
    font-size:12px;
    font-weight:bold;
    cursor: pointer;
}

QComboBox::drop-down {
    subcontrol-origin: padding;
    subcontrol-position: top right;
    width: 1.3em;
    border-left: 0px solid #777;
    border-radius: 0.25em;
}

QComboBox::drop-down::icon {
    image: url('E:/pythonProject_moullin-application.3.5/images/down-arroww.png');
}

        QComboBox:on
        {
            padding-top: 0px;
            padding-left: 0px;        
            selection-background-color: #e4f0f1;
        }
        QComboBox QAbstractItemView
        {
            background-color: #ffffff;
            border-radius: 2px;
            border: 1px solid #76797C;
            color:#000000;
            selection-background-color: #000000;
        }
                                                         
            QComboBox:hover,QDoubleSpinBox:Hover,QPushButton:hover,QAbstractSpinBox:hover,QLineEdit:hover,QTextEdit:hover,QPlainTextEdit:hover,QAbstractView:hover,QTreeView:hover
            {
                border: 1px solid #ff8c00;
                color: #eff0f1;
            }

            


            QLabel
            {
                border: 2px solid black;
            }

            QTabWidget{
                border: 0px transparent black;
            }

            QTabWidget::pane {
                border: 1px solid #76797C;
                padding: 5px;
                margin: 0px;
            }

            QTabBar
            {
                qproperty-drawBase: 0;
                left: 5px; /* move to the right by 5px */
                border-radius: 3px;
            }

            QTabBar:focus
            {
                border: 0px transparent black;
            }

            QTabBar::close-button  {
                image: url(:/qss_icons/Dark_rc/close.png);
                background: transparent;
            }

            QTabBar::close-button:hover
            {
                image: url(:/qss_icons/Dark_rc/close-hover.png);
                background: transparent;
            }

            QTabBar::close-button:pressed {
                image: url(:/qss_icons/Dark_rc/close-pressed.png);
                background: transparent;
            }

            /* TOP TABS */
            QTabBar::tab:top {
                color: #eff0f1;
                border: 1px solid #76797C;
                border-bottom: 1px transparent black;
                background-color: #31363b;
                padding: 5px;
                min-width: 50px;
                border-top-left-radius: 2px;
                border-top-right-radius: 2px;
            }

            QTabBar::tab:top:!selected
            {
                color: #eff0f1;
                background-color: #54575B;
                border: 1px solid #76797C;
                border-bottom: 1px transparent black;
                border-top-left-radius: 2px;
                border-top-right-radius: 2px;    
            }

            QTabBar::tab:top:!selected:hover {
                background-color: #3daee9;
            }

            /* BOTTOM TABS */
            QTabBar::tab:bottom {
                color: #eff0f1;
                border: 1px solid #76797C;
                border-top: 1px transparent black;
                background-color: #31363b;
                padding: 5px;
                border-bottom-left-radius: 2px;
                border-bottom-right-radius: 2px;
                min-width: 50px;
            }

            QTabBar::tab:bottom:!selected
            {
                color: #eff0f1;
                background-color: #54575B;
                border: 1px solid #76797C;
                border-top: 1px transparent black;
                border-bottom-left-radius: 2px;
                border-bottom-right-radius: 2px;
            }

            QTabBar::tab:bottom:!selected:hover {
                background-color: #3daee9;
            }

            /* LEFT TABS */
            QTabBar::tab:left {
                color: #eff0f1;
                border: 1px solid #76797C;
                border-left: 1px transparent black;
                background-color: #31363b;
                padding: 5px;
                border-top-right-radius: 2px;
                border-bottom-right-radius: 2px;
                min-height: 50px;
            }

            QTabBar::tab:left:!selected
            {
                color: #eff0f1;
                background-color: #54575B;
                border: 1px solid #76797C;
                border-left: 1px transparent black;
                border-top-right-radius: 2px;
                border-bottom-right-radius: 2px;
            }

            QTabBar::tab:left:!selected:hover {
                background-color: #3daee9;
            }


            /* RIGHT TABS */
            QTabBar::tab:right {
                color: #eff0f1;
                border: 1px solid #76797C;
                border-right: 1px transparent black;
                background-color: #31363b;
                padding: 5px;
                border-top-left-radius: 2px;
                border-bottom-left-radius: 2px;
                min-height: 50px;
            }

            QTabBar::tab:right:!selected
            {
                color: #eff0f1;
                background-color: #54575B;
                border: 1px solid #76797C;
                border-right: 1px transparent black;
                border-top-left-radius: 2px;
                border-bottom-left-radius: 2px;
            }

            QTabBar::tab:right:!selected:hover {
                background-color: #3daee9;
            }

            QTabBar QToolButton::right-arrow:enabled {
                 image: url(:/qss_icons/Dark_rc/right_arrow.png);
             }

             QTabBar QToolButton::left-arrow:enabled {
                 image: url(:/qss_icons/Dark_rc/left_arrow.png);
             }

            QTabBar QToolButton::right-arrow:disabled {
                 image: url(:/qss_icons/Dark_rc/right_arrow_disabled.png);
             }

             QTabBar QToolButton::left-arrow:disabled {
                 image: url(:/qss_icons/Dark_rc/left_arrow_disabled.png);
             }


            QDockWidget {
                background: #31363b;
                border: 1px solid #403F3F;
                titlebar-close-icon: url(:/qss_icons/Dark_rc/close.png);
                titlebar-normal-icon: url(:/qss_icons/Dark_rc/undock.png);
            }

            QDockWidget::close-button, QDockWidget::float-button {
                border: 1px solid transparent;
                border-radius: 2px;
                background: transparent;
            }

            QDockWidget::close-button:hover, QDockWidget::float-button:hover {
                background: rgba(255, 255, 255, 10);
            }

            QDockWidget::close-button:pressed, QDockWidget::float-button:pressed {
                padding: 1px -1px -1px 1px;
                background: rgba(255, 255, 255, 10);
            }


            QSlider::groove:horizontal {
                border: 1px solid #565a5e;
                height: 4px;
                background: #565a5e;
                margin: 0px;
                border-radius: 2px;
            }

            QSlider::handle:horizontal {
                background: #232629;
                border: 1px solid #565a5e;
                width: 16px;
                height: 16px;
                margin: -8px 0;
                border-radius: 9px;
            }

            QSlider::groove:vertical {
                border: 1px solid #565a5e;
                width: 4px;
                background: #565a5e;
                margin: 0px;
                border-radius: 3px;
            }

            QSlider::handle:vertical {
                background: #232629;
                border: 1px solid #565a5e;
                width: 16px;
                height: 16px;
                margin: 0 -8px;
                border-radius: 9px;
            }

            QToolButton {
                background-color: transparent;
                border: 1px transparent #76797C;
                border-radius: 2px;
                margin: 3px;
                padding: 5px;
            }

            QToolButton[popupMode="1"] { /* only for MenuButtonPopup */
             padding-right: 20px; /* make way for the popup button */
             border: 1px #76797C;
             border-radius: 5px;
            }

            QToolButton[popupMode="2"] { /* only for InstantPopup */
             padding-right: 10px; /* make way for the popup button */
             border: 1px #76797C;
            }


            QToolButton:hover, QToolButton::menu-button:hover {
                background-color: transparent;
                border: 1px solid #3daee9;
                padding: 5px;
            }

            QToolButton:checked, QToolButton:pressed,
                    QToolButton::menu-button:pressed {
                background-color: #3daee9;
                border: 1px solid #3daee9;
                padding: 5px;
            }

            /* the subcontrol below is used only in the InstantPopup or DelayedPopup mode */
            QToolButton::menu-indicator {
                background-color:ff8c00;
                top: -7px; left: -2px; /* shift it a bit */
            }

            /* the subcontrols below are used only in the MenuButtonPopup mode */
            QToolButton::menu-button {
                border: 1px transparent #76797C;
                border-top-right-radius: 6px;
                border-bottom-right-radius: 6px;
                /* 16px width + 4px for border = 20px allocated above */
                width: 16px;
                outline: none;
            }

            QToolButton::menu-arrow {
               background-color:ff8c00;
            }

            QToolButton::menu-arrow:open {
                border: 1px solid #76797C;
            }

            QPushButton::menu-indicator  {
                subcontrol-origin: padding;
                subcontrol-position: bottom right;
                left: 8px;
            }

            QTableView
            {
                border: 1px solid #76797C;
                gridline-color: #31363b;
                background-color: #232629;
            }


            QTableView, QHeaderView
            {
                border-radius: 0px;
            }

            QTableView::item:pressed, QListView::item:pressed, QTreeView::item:pressed  {
                background: #3daee9;
                color: #eff0f1;
            }

            QTableView::item:selected:active, QTreeView::item:selected:active, QListView::item:selected:active  {
                background: #3daee9;
                color: #eff0f1;
            }


            QHeaderView
            {
                background-color: #31363b;
                border: 1px transparent;
                border-radius: 0px;
                margin: 0px;
                padding: 0px;

            }

            QHeaderView::section  {
                background-color: #31363b;
                color: #eff0f1;
                padding: 5px;
                border: 1px solid #76797C;
                border-radius: 0px;
                text-align: center;
            }

            QHeaderView::section::vertical::first, QHeaderView::section::vertical::only-one
            {
                border-top: 1px solid #76797C;
            }

            QHeaderView::section::vertical
            {
                border-top: transparent;
            }

            QHeaderView::section::horizontal::first, QHeaderView::section::horizontal::only-one
            {
                border-left: 1px solid #76797C;
            }

            QHeaderView::section::horizontal
            {
                border-left: transparent;
            }


            QHeaderView::section:checked
             {
                color: white;
                background-color: #334e5e;
             }

             /* style the sort indicator */
            QHeaderView::down-arrow {
                image: url(:/qss_icons/Dark_rc/down_arrow.png);
            }

            QHeaderView::up-arrow {
                image: url(:/qss_icons/Dark_rc/up_arrow.png);
            }


            QTableCornerButton::section {
                background-color: #31363b;
                border: 1px transparent #76797C;
                border-radius: 0px;
            }

            QToolBox  {
                padding: 5px;
                border: 1px transparent black;
            }

            QToolBox::tab {
                color: #eff0f1;
                background-color: #31363b;
                border: 1px solid #76797C;
                border-bottom: 1px transparent #31363b;
                border-top-left-radius: 5px;
                border-top-right-radius: 5px;
            }

            QToolBox::tab:selected { /* italicize selected tabs */
                font: italic;
                background-color: #31363b;
                border-color: #3daee9;
             }

            QStatusBar::item {
                border: 0px transparent dark;
             }


            QFrame[height="3"], QFrame[width="3"] {
                background-color: #76797C;
            }




            QDateEdit
            {
                selection-background-color:#31363b;
                border-style: solid;
                border: 1px solid #76797C;
                border-radius: 2px;
                padding: 1px;
                min-width: 75px;
            }

            QDateEdit:on
            {
                padding-top: 2px;
                padding-left: 2px;
                selection-background-color: #4a4a4a;
            }

            QDateEdit QAbstractItemView
            {
                background-color: #ff8c00;
                border-radius: 2px;
                border: 1px solid #3375A3;
                selection-background-color:ff8c00;
            }

            QDateEdit::drop-down
            {
                subcontrol-origin: padding;
                subcontrol-position: top right;
                width: 15px;
                border-left-width: 0px;
                border-left-color: darkgray;
                border-left-style: solid;
                border-top-right-radius: 3px;
                border-bottom-right-radius: 3px;
            }""")
            self.centralwidget = QtWidgets.QWidget(MainWindow)
            self.centralwidget.setObjectName("centralwidget")
            self.horizontalLayout = QtWidgets.QHBoxLayout(self.centralwidget)
            self.horizontalLayout.setObjectName("horizontalLayout")
            self.horizontalLayout.setContentsMargins(0, 0, 0, 0)
            self.moulinwidget = QtWidgets.QTabWidget(self.centralwidget)
            self.moulinwidget.setObjectName("moulinwidget")

            self.bltendretab = QtWidgets.QWidget()
            self.bltendretab.setStyleSheet("""QToolTip
    {
        border: 1px solid #76797C;
        background-color:  #fff8b0;
        color: white;
        padding: 5px;
        opacity: 200;
    }

    QWidget
    {
        color: #000000;
        background-color:  #D8F9DB;
        selection-background-color:#3daee9;
        selection-color: #eff0f1;
        background-clip: border;
        border-image: none;
        border: 0px transparent black;
        outline: 0;
    }

    QWidget:item:hover
    {
        background-color: #3daee9;
        color: #eff0f1;
    }

    QWidget:item:selected
    {
        background-color: #3daee9;
    }



    QWidget:disabled
    {
        color: #454545;
        background-color: #31363b;
    }

    QAbstractItemView
    {
        alternate-background-color: #31363b;
        color: #eff0f1;
        border: 1px solid 3A3939;
        border-radius: 2px;
    }

    QWidget:focus, QMenuBar:focus
    {
        border: 1px solid #3daee9;
    }

    QTabWidget:focus, QCheckBox:focus, QRadioButton:focus, QSlider:focus
    {
        border: none;
    }

    QLineEdit
    {
        background-color: #ffffff;
        padding: 1px;
        border-style: solid;
        border: 1px solid #000000;
        border-radius: 2px;
        color: #000000;
        font-size:12px;
        font-weight:bold;
    }
    QDoubleSpinBox
    {
        background-color: #ffffff;
        padding: 1px;
        border-style: solid;
        border: 2px solid #76797C;
        border-radius: 4px;
        border-color: #000000;
        color:#000000;
        font-size:12px;
        font-weight:bold;

    }

    QDoubleSpinBox:focus{
        background-color: #f2f2f2;
        border-style: solid;
        border: 2px solid #76797C;
        border-radius: 4px;
        border-color: #ff8c00;
    }

    QDoubleSpinBox::drop-down
    {
        subcontrol-origin: padding;
        subcontrol-position: top right;
        width: 1px;

        border-left-width: 0px;
        border-left-color: #232629;
        border-left-style: solid;
        border-top-right-radius: 1px;
        border-bottom-right-radius: 1px;
    }



    QGroupBox {
        border:1px solid #76797C;
        border-radius: 2px;
        margin-top: 20px;
    }

    QGroupBox::title {
        subcontrol-origin: margin;
        subcontrol-position: top center;
        padding-left: 10px;
        padding-right: 10px;
        padding-top: 10px;
    }

    QAbstractScrollArea
    {
        border-radius: 2px;
        border: 1px solid #76797C;
        background-color: transparent;
    }

    QScrollBar:horizontal
    {
        height: 15px;
        margin: 3px 15px 3px 15px;
        border: 1px transparent #2A2929;
        border-radius: 4px;
        background-color: #2A2929;
    }

    QScrollBar::handle:horizontal
    {
        background-color: #605F5F;
        min-width: 5px;
        border-radius: 4px;
    }

    QScrollBar::add-line:horizontal
    {
        margin: 0px 3px 0px 3px;
        border-image: url(:/qss_icons/Dark_rc/right_arrow_disabled.png);
        width: 10px;
        height: 10px;
        subcontrol-position: right;
        subcontrol-origin: margin;
    }

    QScrollBar::sub-line:horizontal
    {
        margin: 0px 3px 0px 3px;
        border-image: url(:/qss_icons/Dark_rc/left_arrow_disabled.png);
        height: 10px;
        width: 10px;
        subcontrol-position: left;
        subcontrol-origin: margin;
    }

    QScrollBar::add-line:horizontal:hover,QScrollBar::add-line:horizontal:on
    {
        border-image: url(:/qss_icons/Dark_rc/right_arrow.png);
        height: 10px;
        width: 10px;
        subcontrol-position: right;
        subcontrol-origin: margin;
    }


    QScrollBar::sub-line:horizontal:hover, QScrollBar::sub-line:horizontal:on
    {
        border-image: url(:/qss_icons/Dark_rc/left_arrow.png);
        height: 10px;
        width: 10px;
        subcontrol-position: left;
        subcontrol-origin: margin;
    }

    QScrollBar::up-arrow:horizontal, QScrollBar::down-arrow:horizontal
    {
        background: none;
    }


    QScrollBar::add-page:horizontal, QScrollBar::sub-page:horizontal
    {
        background: none;
    }

    QScrollBar:vertical
    {
        background-color: #2A2929;
        width: 15px;
        margin: 15px 3px 15px 3px;
        border: 1px transparent #2A2929;
        border-radius: 4px;
    }

    QScrollBar::handle:vertical
    {
        background-color: #605F5F;
        min-height: 5px;
        border-radius: 4px;
    }

    QScrollBar::sub-line:vertical
    {
        margin: 3px 0px 3px 0px;
        border-image: url(:/qss_icons/Dark_rc/up_arrow_disabled.png);
        height: 10px;
        width: 10px;
        subcontrol-position: top;
        subcontrol-origin: margin;
    }

    QScrollBar::add-line:vertical
    {
        margin: 3px 0px 3px 0px;
        border-image: url(:/qss_icons/Dark_rc/down_arrow_disabled.png);
        height: 10px;
        width: 10px;
        subcontrol-position: bottom;
        subcontrol-origin: margin;
    }

    QScrollBar::sub-line:vertical:hover,QScrollBar::sub-line:vertical:on
    {

        border-image: url(:/qss_icons/Dark_rc/up_arrow.png);
        height: 10px;
        width: 10px;
        subcontrol-position: top;
        subcontrol-origin: margin;
    }


    QScrollBar::add-line:vertical:hover, QScrollBar::add-line:vertical:on
    {
        border-image: url(:/qss_icons/Dark_rc/down_arrow.png);
        height: 10px;
        width: 10px;
        subcontrol-position: bottom;
        subcontrol-origin: margin;
    }

    QScrollBar::up-arrow:vertical, QScrollBar::down-arrow:vertical
    {
        background: none;
    }


    QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical
    {
        background: none;
    }

    QTextEdit
    {
        background-color: #ffffff;
        color: #000000;
        border: 1px solid #76797C;
        font-size:12px;
        font-weight:bold;
    }

    QPlainTextEdit
    {
        background-color: #232629;;
        color: #eff0f1;
        border-radius: 2px;
        border: 1px solid #76797C;
    }

    QHeaderView::section
    {
        background-color: #76797C;
        color: #eff0f1;
        padding: 5px;
        border: 1px solid #76797C;
    }

    QSizeGrip {
        image: url(:/qss_icons/Dark_rc/sizegrip.png);
        width: 12px;
        height: 12px;
    }


    QMainWindow::separator
    {
        background-color: #31363b;
        color: white;
        padding-left: 4px;
        spacing: 2px;
        border: 1px dashed #76797C;
    }

    QMainWindow::separator:hover
    {

        background-color: #787876;
        color: white;
        padding-left: 4px;
        border: 1px solid #76797C;
        spacing: 2px;
    }


    QMenu::separator
    {
        height: 1px;
        background-color: #76797C;
        color: white;
        padding-left: 4px;
        margin-left: 10px;
        margin-right: 5px;
    }


    QFrame
    {
        border-radius: 2px;
        border: 1px solid #76797C;
    }

    QFrame[frameShape="0"]
    {
        border-radius: 2px;
        border: 1px transparent #76797C;
    }

    QStackedWidget
    {
        border: 1px transparent black;
    }


    QPushButton
    {
        color: #000000;
        background-color:#84dbc8;
        border-width: 1px;
        border-color: #1e1e1e;
        border-style: solid;
        border-radius: 6;
        padding: 3px;
        font-size: 12px;
        padding-left: 5px;
        padding-right: 5px;
        min-width: 40px;

    }

    QPushButton:disabled
    {
        background-color: #31363b;
        border-width: 1px;
        border-color: #454545;
        border-style: solid;
        padding-top: 5px;
        padding-bottom: 5px;
        padding-left: 10px;
        padding-right: 10px;
        border-radius: 2px;
        color: #454545;
    }

    QPushButton:focus {
        background-color: #3daee9;
        color: white;
    }

    QPushButton:pressed
    {
        background-color: #3daee9;
        padding-top: -15px;
        padding-bottom: -17px;
    }

    QComboBox {
    background-color: #ffffff;
    border: 1px solid #76797C;
    color:#000000;
    border-radius: 0.25em;
    padding: 0.50em 0.50em;
    font-size:12px;
    font-weight:bold;
    cursor: pointer;
}

QComboBox::drop-down {
    subcontrol-origin: padding;
    subcontrol-position: top right;
    width: 1.3em;
    border-left: 0px solid #777;
    border-radius: 0.25em;
}

QComboBox::drop-down::icon {
    image: url('E:/pythonProject_moullin-application.3.5/images/down-arroww.png');
}

        QComboBox:on
        {
            padding-top: 1px;
            padding-left: 1px;
            selection-background-color: #e4f0f1;
        }
        QComboBox QAbstractItemView
        {
            background-color: #ffffff;
            border-radius: 2px;
            border: 1px solid #76797C;
            color:#000000;
            selection-background-color: #000000;
        }

    QPushButton:checked{
        background-color: #76797C;
        border-color: #6A6969;
    }

    QComboBox:hover,QDoubleSpinBox:Hover,QPushButton:hover,QAbstractSpinBox:hover,QLineEdit:hover,QTextEdit:hover,QPlainTextEdit:hover,QAbstractView:hover,QTreeView:hover
    {
        border: 1px solid #ff8c00;
        color: #000000;
    }

    

    QLabel
    {
        border: 2px solid black;
        font-size:13px;
        font-weight:bold;
    }

    QTabWidget{
        border: 0px transparent black;
    }

    QTabWidget::pane {
        border: 1px solid #76797C;
        padding: 5px;
        margin: 0px;
    }

    QTabBar
    {
        qproperty-drawBase: 0;
        left: 5px; /* move to the right by 5px */
        border-radius: 3px;
    }

    QTabBar:focus
    {
        border: 0px transparent black;
    }

    QTabBar::close-button  {
        image: url(:/qss_icons/Dark_rc/close.png);
        background: transparent;
    }

    QTabBar::close-button:hover
    {
        image: url(:/qss_icons/Dark_rc/close-hover.png);
        background: transparent;
    }

    QTabBar::close-button:pressed {
        image: url(:/qss_icons/Dark_rc/close-pressed.png);
        background: transparent;
    }

    /* TOP TABS */
    QTabBar::tab:top {
        color: #eff0f1;
        border: 1px solid #76797C;
        border-bottom: 1px transparent black;
        background-color: #31363b;
        padding: 5px;
        min-width: 50px;
        border-top-left-radius: 2px;
        border-top-right-radius: 2px;
    }

    QTabBar::tab:top:!selected
    {
        color: #eff0f1;
        background-color: #54575B;
        border: 1px solid #76797C;
        border-bottom: 1px transparent black;
        border-top-left-radius: 2px;
        border-top-right-radius: 2px;    
    }

    QTabBar::tab:top:!selected:hover {
        background-color: #3daee9;
    }

    /* BOTTOM TABS */
    QTabBar::tab:bottom {
        color: #eff0f1;
        border: 1px solid #76797C;
        border-top: 1px transparent black;
        background-color: #31363b;
        padding: 5px;
        border-bottom-left-radius: 2px;
        border-bottom-right-radius: 2px;
        min-width: 50px;
    }

    QTabBar::tab:bottom:!selected
    {
        color: #eff0f1;
        background-color: #54575B;
        border: 1px solid #76797C;
        border-top: 1px transparent black;
        border-bottom-left-radius: 2px;
        border-bottom-right-radius: 2px;
    }

    QTabBar::tab:bottom:!selected:hover {
        background-color: #3daee9;
    }

    /* LEFT TABS */
    QTabBar::tab:left {
        color: #eff0f1;
        border: 1px solid #76797C;
        border-left: 1px transparent black;
        background-color: #31363b;
        padding: 5px;
        border-top-right-radius: 2px;
        border-bottom-right-radius: 2px;
        min-height: 50px;
    }

    QTabBar::tab:left:!selected
    {
        color: #eff0f1;
        background-color: #54575B;
        border: 1px solid #76797C;
        border-left: 1px transparent black;
        border-top-right-radius: 2px;
        border-bottom-right-radius: 2px;
    }

    QTabBar::tab:left:!selected:hover {
        background-color: #3daee9;
    }


    /* RIGHT TABS */
    QTabBar::tab:right {
        color: #eff0f1;
        border: 1px solid #76797C;
        border-right: 1px transparent black;
        background-color: #31363b;
        padding: 5px;
        border-top-left-radius: 2px;
        border-bottom-left-radius: 2px;
        min-height: 50px;
    }

    QTabBar::tab:right:!selected
    {
        color: #eff0f1;
        background-color: #54575B;
        border: 1px solid #76797C;
        border-right: 1px transparent black;
        border-top-left-radius: 2px;
        border-bottom-left-radius: 2px;
    }

    QTabBar::tab:right:!selected:hover {
        background-color: #3daee9;
    }

    QTabBar QToolButton::right-arrow:enabled {
         image: url(:/qss_icons/Dark_rc/right_arrow.png);
     }

     QTabBar QToolButton::left-arrow:enabled {
         image: url(:/qss_icons/Dark_rc/left_arrow.png);
     }

    QTabBar QToolButton::right-arrow:disabled {
         image: url(:/qss_icons/Dark_rc/right_arrow_disabled.png);
     }

     QTabBar QToolButton::left-arrow:disabled {
         image: url(:/qss_icons/Dark_rc/left_arrow_disabled.png);
     }


    QDockWidget {
        background: #31363b;
        border: 1px solid #403F3F;
        titlebar-close-icon: url(:/qss_icons/Dark_rc/close.png);
        titlebar-normal-icon: url(:/qss_icons/Dark_rc/undock.png);
    }

    QDockWidget::close-button, QDockWidget::float-button {
        border: 1px solid transparent;
        border-radius: 2px;
        background: transparent;
    }

    QDockWidget::close-button:hover, QDockWidget::float-button:hover {
        background: rgba(255, 255, 255, 10);
    }

    QDockWidget::close-button:pressed, QDockWidget::float-button:pressed {
        padding: 1px -1px -1px 1px;
        background: rgba(255, 255, 255, 10);
    }


    QSlider::groove:horizontal {
        border: 1px solid #565a5e;
        height: 4px;
        background: #565a5e;
        margin: 0px;
        border-radius: 2px;
    }

    QSlider::handle:horizontal {
        background: #232629;
        border: 1px solid #565a5e;
        width: 16px;
        height: 16px;
        margin: -8px 0;
        border-radius: 9px;
    }

    QSlider::groove:vertical {
        border: 1px solid #565a5e;
        width: 4px;
        background: #565a5e;
        margin: 0px;
        border-radius: 3px;
    }

    QSlider::handle:vertical {
        background: #232629;
        border: 1px solid #565a5e;
        width: 16px;
        height: 16px;
        margin: 0 -8px;
        border-radius: 9px;
    }

    QToolButton {
        background-color: transparent;
        border: 1px transparent #76797C;
        border-radius: 2px;
        margin: 3px;
        padding: 5px;
    }

    QToolButton[popupMode="1"] { /* only for MenuButtonPopup */
     padding-right: 20px; /* make way for the popup button */
     border: 1px #76797C;
     border-radius: 5px;
    }

    QToolButton[popupMode="2"] { /* only for InstantPopup */
     padding-right: 10px; /* make way for the popup button */
     border: 1px #76797C;
    }


    QToolButton:hover, QToolButton::menu-button:hover {
        background-color: transparent;
        border: 1px solid #3daee9;
        padding: 5px;
    }

    QToolButton:checked, QToolButton:pressed,
            QToolButton::menu-button:pressed {
        background-color: #3daee9;
        border: 1px solid #3daee9;
        padding: 5px;
    }

    /* the subcontrol below is used only in the InstantPopup or DelayedPopup mode */
    QToolButton::menu-indicator {
        background-color:ff8c00;
        top: -7px; left: -2px; /* shift it a bit */
    }

    /* the subcontrols below are used only in the MenuButtonPopup mode */
    QToolButton::menu-button {
        border: 1px transparent #76797C;
        border-top-right-radius: 6px;
        border-bottom-right-radius: 6px;
        /* 16px width + 4px for border = 20px allocated above */
        width: 16px;
        outline: none;
    }

    QToolButton::menu-arrow {
       background-color:ff8c00;
    }

    QToolButton::menu-arrow:open {
        border: 1px solid #76797C;
    }

    QPushButton::menu-indicator  {
        subcontrol-origin: padding;
        subcontrol-position: bottom right;
        left: 8px;
    }

    QTableView
    {
        border: 1px solid #76797C;
        gridline-color: #31363b;
        background-color: #232629;
    }


    QTableView, QHeaderView
    {
        border-radius: 0px;
    }

    QTableView::item:pressed, QListView::item:pressed, QTreeView::item:pressed  {
        background: #3daee9;
        color: #eff0f1;
    }

    QTableView::item:selected:active, QTreeView::item:selected:active, QListView::item:selected:active  {
        background: #3daee9;
        color: #eff0f1;
    }


    QHeaderView
    {
        background-color: #31363b;
        border: 1px transparent;
        border-radius: 0px;
        margin: 0px;
        padding: 0px;

    }

    QHeaderView::section  {
        background-color: #31363b;
        color: #eff0f1;
        padding: 5px;
        border: 1px solid #76797C;
        border-radius: 0px;
        text-align: center;
    }

    QHeaderView::section::vertical::first, QHeaderView::section::vertical::only-one
    {
        border-top: 1px solid #76797C;
    }

    QHeaderView::section::vertical
    {
        border-top: transparent;
    }

    QHeaderView::section::horizontal::first, QHeaderView::section::horizontal::only-one
    {
        border-left: 1px solid #76797C;
    }

    QHeaderView::section::horizontal
    {
        border-left: transparent;
    }


    QHeaderView::section:checked
     {
        color: white;
        background-color: #334e5e;
     }

     /* style the sort indicator */
    QHeaderView::down-arrow {
        image: url(:/qss_icons/Dark_rc/down_arrow.png);
    }

    QHeaderView::up-arrow {
        image: url(:/qss_icons/Dark_rc/up_arrow.png);
    }


    QTableCornerButton::section {
        background-color: #31363b;
        border: 1px transparent #76797C;
        border-radius: 0px;
    }

    QToolBox  {
        padding: 5px;
        border: 1px transparent black;
    }

    QToolBox::tab {
        color: #eff0f1;
        background-color: #31363b;
        border: 1px solid #76797C;
        border-bottom: 1px transparent #31363b;
        border-top-left-radius: 5px;
        border-top-right-radius: 5px;
    }

    QToolBox::tab:selected { /* italicize selected tabs */
        font: italic;
        background-color: #31363b;
        border-color: #3daee9;
     }

    QStatusBar::item {
        border: 0px transparent dark;
     }


    QFrame[height="3"], QFrame[width="3"] {
        background-color: #76797C;
    }




    QDateEdit
    {
        selection-background-color:#ffffff;
        border-style: solid;
        border: 1px solid #000000;
        border-radius: 2px;
        padding: 1px;
        min-width: 75px;
    }

    QDateEdit:on
    {
        padding-top: 2px;
        padding-left: 2px;
        selection-background-color: #ffffff;
    }

    QDateEdit QAbstractItemView
    {
        background-color: #ffffff;
        border-radius: 2px;
        border: 1px solid #3375A3;
        selection-background-color:000000;
    }

    QDateEdit::drop-down
    {
        subcontrol-origin: padding;
        subcontrol-position: top right;
        width: 15px;
        border-left-width: 0px;
        border-left-color: #000000;
        border-left-style: solid;
        border-top-right-radius: 3px;
        border-bottom-right-radius: 3px;
    }""")
            self.bltendretab.setObjectName("bltendretab")
            self.moulinwidget.addTab(self.bltendretab, "")

            self.font = QtGui.QFont()
            self.font.setBold(True)
            self.font.setPointSize(10)
            ##########text bul# ettin######
            self.paramétre = QtWidgets.QLabel("Paramètre", self.bltendretab)
            self.paramétre.move(30, 145)
            self.paramétre.resize(80, 20)
            self.paramétre.setFont(self.font)

            self.txtpsfont = QtGui.QFont()
            self.txtpsfont.setBold(True)
            self.txtpsfont.setPointSize(9)
            ################Limites(sans bon ni réf)###############
            self.valeur = QtWidgets.QLabel("""Limite ssans bon ni réf)""", self.bltendretab)
            self.valeur.move(165, 127)
            self.valeur.resize(200, 55)
            self.valeur.setFont(self.font)
            ######################Limites(sans bon ni réf)################
            self.ps = QtWidgets.QLineEdit("Poids spécifique (kg/hl):   (75.500-75.899)", self.bltendretab,
                                          readOnly=True)
            self.ps.resize(319, 20)
            self.ps.move(30, 167)
            self.ps.setFont(self.txtpsfont)
            self.ps.setStyleSheet(" border: 2px solid bleu ;border-radius: 4px;padding: 0px")
            ###############################humidite#############
            self.humidite = QtWidgets.QLineEdit("Teneur en eau(%): (13.5-15)", self.bltendretab, readOnly=True)
            self.humidite.resize(319, 20)
            self.humidite.move(30, 188)
            self.humidite.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            self.humidite.setFont(self.txtpsfont)

            #######################ergot#########################
            self.ergot = QtWidgets.QLineEdit("Ergo(% :\t<0.001 ", self.bltendretab, readOnly=True)
            self.ergot.resize(319, 20)
            self.ergot.move(30, 209)
            self.ergot.setStyleSheet("background-color: #232629")
            self.ergot.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            self.ergot.setFont(self.txtpsfont)

            #########################Graines nuisibles (%)##########
            self.grainnuisible = QtWidgets.QLineEdit("Graines nuisibles(%):\t<0.001", self.bltendretab, readOnly=True)
            self.grainnuisible.resize(319, 20)
            self.grainnuisible.move(30, 230)
            self.grainnuisible.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            self.grainnuisible.setFont(self.txtpsfont)
            #############################Débris végétaux (%)########
            self.débrisvé = QtWidgets.QLineEdit("Débris végétaux(%):     ", self.bltendretab, readOnly=True)
            self.débrisvé.resize(319, 20)
            self.débrisvé.move(30, 251)
            self.débrisvé.setFont(self.txtpsfont)
            self.débrisvé.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            #########################Matière inerte (%)################
            self.matiéreinrt = QtWidgets.QLineEdit("Matière inerte(%):", self.bltendretab, readOnly=True)
            self.matiéreinrt.resize(319, 20)
            self.matiéreinrt.move(30, 272)
            self.matiéreinrt.setFont(self.txtpsfont)
            self.matiéreinrt.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            ################################Grains chauffés (%)############################
            self.grainchaufé = QtWidgets.QLineEdit("Grains chauffés(%):    ", self.bltendretab, readOnly=True)
            self.grainchaufé.resize(319, 20)
            self.grainchaufé.move(30, 293)
            self.grainchaufé.setFont(self.txtpsfont)
            self.grainchaufé.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            ########################################Grains sans valeur (%)#######################################
            self.grainsanvaleur = QtWidgets.QLineEdit("Grains sans valeur(%):", self.bltendretab, readOnly=True)
            self.grainsanvaleur.resize(319, 20)
            self.grainsanvaleur.move(30, 314)
            self.grainsanvaleur.setFont(self.txtpsfont)
            self.grainsanvaleur.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            #######################################Grains cariés##########################################
            self.graincarré = QtWidgets.QLineEdit("Grains cariés:   ", self.bltendretab, readOnly=True)
            self.graincarré.resize(319, 20)
            self.graincarré.move(30, 335)
            self.graincarré.setFont(self.txtpsfont)
            self.graincarré.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")

            #######################################Total(%) 1er#####################################################
            self.totalprem = QtWidgets.QLineEdit("TOTAL 1ére CAT:     ≤1", self.bltendretab, readOnly=True)
            self.totalprem.resize(319, 20)
            self.totalprem.move(30, 356)
            self.totalprem.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            self.totalprem.setFont(self.txtpsfont)
            ##############################################Grains cassés (%) #########################################################
            self.graincassé = QtWidgets.QLineEdit("Grains cassés(%):   ≤2", self.bltendretab, readOnly=True)
            self.graincassé.move(30, 377)
            self.graincassé.resize(319, 20)
            self.graincassé.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            self.graincassé.setFont(self.txtpsfont)
            #########################################################Gains échaudés (%)#####################################################
            self.grainechaude = QtWidgets.QLineEdit("Gains échaudés(%):   ", self.bltendretab, readOnly=True)
            self.grainechaude.move(30, 419)
            self.grainechaude.resize(319, 20)
            self.grainechaude.setFont(self.txtpsfont)
            self.grainechaude.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            #####################################################Grains maigres (%)########################################################
            self.grainmaigre = QtWidgets.QLineEdit("Grains maigres(%):", self.bltendretab, readOnly=True)
            self.grainmaigre.move(30, 398)
            self.grainmaigre.resize(319, 20)
            self.grainmaigre.setFont(self.txtpsfont)
            self.grainmaigre.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            ##########################################################Grains germés (%)###################################################
            self.graigermé = QtWidgets.QLineEdit("Grains germés(%): ≤2", self.bltendretab, readOnly=True)
            self.graigermé.move(30, 440)
            self.graigermé.resize(319, 20)
            self.graigermé.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            self.graigermé.setFont(self.txtpsfont)
            ##########################################################Grain punaisés (%)#########################################################
            self.grainpunaisé = QtWidgets.QLineEdit("Grain punaisés(%): ≤1", self.bltendretab, readOnly=True)
            self.grainpunaisé.move(30, 461)
            self.grainpunaisé.resize(319, 20)
            self.grainpunaisé.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            self.grainpunaisé.setFont(self.txtpsfont)

            #######################################################################Grains piqués (%)##########################################
            self.grainpiqué = QtWidgets.QLineEdit("Grains piqués(%):  ", self.bltendretab, readOnly=True)
            self.grainpiqué.move(30, 482)
            self.grainpiqué.resize(319, 20)
            self.grainpiqué.setFont(self.txtpsfont)
            self.grainpiqué.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            #######################################################################Grains boutés « faible » (%)#######################################
            self.grainboutef = QtWidgets.QLineEdit("Grains boutés « faible » (%):", self.bltendretab, readOnly=True)
            self.grainboutef.move(30, 503)
            self.grainboutef.resize(319, 20)
            self.grainboutef.setFont(self.txtpsfont)
            self.grainboutef.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            ####################################################################Grains boutés  « forte » (%)######################################
            self.grainbouté = QtWidgets.QLineEdit("Grains boutés  « forte » (%):", self.bltendretab, readOnly=True)
            self.grainbouté.move(30, 524)
            self.grainbouté.resize(319, 20)
            self.grainbouté.setFont(self.txtpsfont)
            self.grainbouté.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            ##################################################Grains mouchetés (%)########################################################
            self.grainmouchté = QtWidgets.QLineEdit("Grains mouchetés (%):", self.bltendretab, readOnly=True)
            self.grainmouchté.move(30, 545)
            self.grainmouchté.resize(319, 20)
            self.grainmouchté.setFont(self.txtpsfont)
            self.grainmouchté.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            ########################################################Grain étrangers Utilisables pour le bétail (%)##################################################
            self.grainetrangé = QtWidgets.QLineEdit("Grain étrangers Utilisables pour le bétail (%):  ",
                                                    self.bltendretab, readOnly=True)
            self.grainetrangé.move(30, 566)
            self.grainetrangé.resize(319, 20)
            self.grainetrangé.setFont(self.txtpsfont)
            self.grainetrangé.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            ######################################################################Total(%)######################################
            self.totaldem = QtWidgets.QLineEdit("Total(%)  Imp2eme cat   ≤5", self.bltendretab, readOnly=True)
            self.totaldem.move(30, 587)
            self.totaldem.resize(319, 20)
            self.totaldem.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            self.totaldem.setFont(self.txtpsfont)
            ###########################################################################################################
            self.totalbetr = QtWidgets.QLineEdit("\tTotal des Bonifications et Réfactions", self.bltendretab,
                                                 readOnly=True)
            self.totalbetr.move(30, 608)
            self.totalbetr.resize(419, 25)
            self.totalbetr.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            self.totalbetr.setFont(self.txtpsfont)
            #################label valeure##############
            self.valeur = QtWidgets.QLabel("valeur", self.bltendretab)
            self.valeur.move(350, 144)
            self.valeur.resize(100, 20)
            self.valeur.setFont(self.font)
            ######################Limites(sans bon ni réf)################
            self.vps = QtWidgets.QDoubleSpinBox(self.bltendretab)
            self.vps.setRange(60, 81.00)
            self.vps.setSpecialValueText(' ')
            self.vps.resize(100, 20)
            self.vps.move(350, 167)
            self.vps.setFont(self.txtpsfont)
            # self.vps.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            ###############################humidite#############
            self.vhumidite = QtWidgets.QDoubleSpinBox(self.bltendretab)
            self.vhumidite.setRange(4, 18)
            self.vhumidite.resize(100, 20)
            self.vhumidite.setSpecialValueText(' ')
            self.vhumidite.move(350, 188)
            self.vhumidite.setFont(self.txtpsfont)
            # self.vhumidite.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            #######################ergot#########################
            self.vergot = QtWidgets.QDoubleSpinBox(self.bltendretab, readOnly=False)
            self.vergot.setRange(0, 1)
            self.vergot.setSpecialValueText(' ')
            self.vergot.resize(100, 20)
            self.vergot.move(350, 209)
            # self.vergot.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            #########################Graines nuisibles (%)##########
            self.vgrainnuisible = QtWidgets.QDoubleSpinBox(self.bltendretab, readOnly=False)
            self.vgrainnuisible.setRange(0, 1)
            self.vgrainnuisible.setSpecialValueText(' ')
            self.vgrainnuisible.resize(100, 20)
            self.vgrainnuisible.move(350, 230)
            self.vgrainnuisible.setFont(self.txtpsfont)
            # self.vgrainnuisible.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            #############################Débris végétaux (%)########
            self.vdébrisvé = QtWidgets.QDoubleSpinBox(self.bltendretab, readOnly=False)
            self.vdébrisvé.setRange(0, 10)
            self.vdébrisvé.setSpecialValueText(' ')
            self.vdébrisvé.resize(100, 20)
            self.vdébrisvé.move(350, 251)
            self.vdébrisvé.setFont(self.txtpsfont)
            # self.vdébrisvé.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            #########################Matière inerte (%)################
            self.vmatiéreinrt = QtWidgets.QDoubleSpinBox(self.bltendretab, readOnly=False)
            self.vmatiéreinrt.setRange(0, 10)
            self.vmatiéreinrt.setSpecialValueText(' ')
            self.vmatiéreinrt.resize(100, 20)
            self.vmatiéreinrt.move(350, 272)
            self.vmatiéreinrt.setFont(self.txtpsfont)
            # self.vmatiéreinrt.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            ################################Grains chauffés (%)############################
            self.vgrainchaufé = QtWidgets.QDoubleSpinBox(self.bltendretab, readOnly=False)
            self.vgrainchaufé.setRange(0, 7)
            self.vgrainchaufé.setSpecialValueText(' ')
            self.vgrainchaufé.resize(100, 20)
            self.vgrainchaufé.move(350, 293)
            self.vgrainchaufé.setFont(self.txtpsfont)
            # self.vgrainchaufé.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            ########################################Grains sans valeur (%)#######################################
            self.vgrainsanvaleur = QtWidgets.QDoubleSpinBox(self.bltendretab, readOnly=False)
            self.vgrainsanvaleur.setSpecialValueText(' ')
            self.vgrainsanvaleur.setRange(0, 10)
            self.vgrainsanvaleur.resize(100, 20)
            self.vgrainsanvaleur.move(350, 314)
            self.vgrainsanvaleur.setFont(self.txtpsfont)
            # self.vgrainsanvaleur.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            #######################################Grains cariés##########################################
            self.vgraincarré = QtWidgets.QDoubleSpinBox(self.bltendretab, readOnly=False)
            self.vgraincarré.setSpecialValueText(' ')
            self.vgraincarré.setFont(self.txtpsfont)
            self.vgraincarré.setRange(0, 10)
            self.vgraincarré.resize(100, 20)
            self.vgraincarré.move(350, 335)
            # self.vgraincarré.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            #######################################Total(%) 1er#####################################################
            self.vtotalprem = QtWidgets.QDoubleSpinBox(self.bltendretab, readOnly=True)
            self.vtotalprem.setSpecialValueText(' ')
            self.vtotalprem.setRange(0, 20)
            self.vtotalprem.resize(100, 20)
            self.vtotalprem.move(350, 356)
            self.vtotalprem.setFont(self.txtpsfont)
            # self.vtotalprem.setStyleSheet("background-color:##c7d3cf;color:#000000;border: 2px solid bleu;border-radius: 4px;padding: 0px")
            ##############################################Grains cassés (%) #########################################################
            self.vgraincassé = QtWidgets.QDoubleSpinBox(self.bltendretab)
            self.vgraincassé.move(350, 377)
            self.vgraincassé.resize(100, 20)
            self.vgraincassé.setRange(0, 20)
            self.vgraincassé.setSpecialValueText(" ")
            self.vgraincassé.setFont(self.txtpsfont)
            # self.vgraincassé.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            #########################################################Gains échaudés (%)#####################################################
            self.vgrainechaude = QtWidgets.QDoubleSpinBox(self.bltendretab)
            self.vgrainechaude.setSpecialValueText(" ")
            self.vgrainechaude.setRange(0, 10)
            self.vgrainechaude.move(350, 419)
            self.vgrainechaude.resize(100, 20)
            self.vgrainechaude.setFont(self.txtpsfont)
            # self.vgrainechaude.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            #####################################################Grains maigres (%)########################################################
            self.vgrainmaigre = QtWidgets.QDoubleSpinBox(self.bltendretab)
            self.vgrainmaigre.setRange(0, 10)
            self.vgrainmaigre.setSpecialValueText(" ")
            self.vgrainmaigre.move(350, 398)
            self.vgrainmaigre.setFont(self.txtpsfont)
            self.vgrainmaigre.resize(100, 20)
            # self.vgrainmaigre.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            ##########################################################Grains germés (%)###################################################
            self.vgraigermé = QtWidgets.QDoubleSpinBox(self.bltendretab, readOnly=False)
            self.vgraigermé.move(350, 440)
            self.vgraigermé.resize(100, 20)
            self.vgraigermé.setSpecialValueText('  ')
            self.vgraigermé.setFont(self.txtpsfont)
            # self.vgraigermé.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            ##########################################################Grain punaisés (%)#########################################################
            self.vgrainpunaisé = QtWidgets.QDoubleSpinBox(self.bltendretab, readOnly=False)
            self.vgrainpunaisé.move(350, 461)
            self.vgrainpunaisé.resize(100, 20)
            self.vgrainpunaisé.setSpecialValueText('   ')
            self.vgrainpunaisé.setFont(self.txtpsfont)
            # self.vgrainpunaisé.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            #######################################################################Grains piqués (%)##########################################
            self.vgrainpiqué = QtWidgets.QDoubleSpinBox(self.bltendretab, readOnly=False)
            self.vgrainpiqué.move(350, 482)
            self.vgrainpiqué.resize(100, 20)
            self.vgrainpiqué.setSpecialValueText('  ')
            self.vgrainpiqué.setFont(self.txtpsfont)
            # self.vgrainpiqué.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            #######################################################################Grains boutés « faible » (%)#######################################
            self.vgrainboutef = QtWidgets.QDoubleSpinBox(self.bltendretab, readOnly=False)
            self.vgrainboutef.move(350, 503)
            self.vgrainboutef.resize(100, 20)
            self.vgrainboutef.setRange(0, 10)
            self.vgrainboutef.setSpecialValueText('  ')
            self.vgrainboutef.setFont(self.txtpsfont)
            # self.vgrainboutef.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            ####################################################################Grains boutés  « forte » (%)######################################
            self.vgrainbouté = QtWidgets.QDoubleSpinBox(self.bltendretab, readOnly=False)
            self.vgrainbouté.move(350, 524)
            self.vgrainbouté.resize(100, 20)
            self.vgrainbouté.setRange(0, 10)
            self.vgrainbouté.setSpecialValueText('  ')
            self.vgrainboutef.setFont(self.txtpsfont)
            # self.vgrainbouté.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            ##################################################Grains mouchetés (%)########################################################
            self.vgrainmouchté = QtWidgets.QDoubleSpinBox(self.bltendretab)
            self.vgrainmouchté.move(350, 545)
            self.vgrainmouchté.resize(100, 20)
            self.vgrainmouchté.setRange(0, 5)
            self.vgrainmouchté.setSpecialValueText(' ')
            self.vgrainmouchté.setFont(self.txtpsfont)
            # self.vgrainmouchté.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            ########################################################Grain étrangers Utilisables pour le bétail (%)##################################################
            self.vgrainetrangé = QtWidgets.QDoubleSpinBox(self.bltendretab)
            self.vgrainetrangé.move(350, 566)
            self.vgrainetrangé.resize(100, 20)
            self.vgrainetrangé.setRange(0, 10)
            self.vgrainetrangé.setSpecialValueText(' ')
            self.vgrainetrangé.setFont(self.txtpsfont)
            # self.vgrainetrangé.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            ######################################################################Total(%)######################################
            self.vtotaldem = QtWidgets.QDoubleSpinBox(self.bltendretab, readOnly=True)
            self.vtotaldem.setRange(1, 30)
            self.vtotaldem.move(350, 587)
            self.vtotaldem.resize(100, 20)
            self.vtotaldem.setSpecialValueText(' ')
            self.vtotaldem.setFont(self.txtpsfont)
            # self.vtotaldem.setStyleSheet("background-color:#ffffff;color:#000000;border: 2px solid bleu;border-radius: 4px;padding: 0px")
            #######################################################réfaction##############################################
            #################label valeure##############
            self.rvaleur = QtWidgets.QLabel("Réfaction", self.bltendretab)
            self.rvaleur.move(530, 145)
            self.rvaleur.resize(80, 20)
            self.rvaleur.setFont(self.font)
            ######################Limites(sans bon ni réf)################
            self.rps = QtWidgets.QDoubleSpinBox(self.bltendretab, readOnly=True)
            self.rps.setSpecialValueText(" ")
            self.rps.setFont(self.txtpsfont)
            self.rps.resize(100, 20)
            self.rps.move(530, 167)
            self.rps.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            ###############################humidite#############
            self.rhumidite = QtWidgets.QDoubleSpinBox(self.bltendretab, readOnly=True)
            self.rhumidite.setSpecialValueText(" ")
            self.rhumidite.setFont(self.txtpsfont)
            self.rhumidite.resize(100, 20)
            self.rhumidite.move(530, 188)
            self.rhumidite.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            #######################ergot#########################
            self.rergot = QtWidgets.QLineEdit("", self.bltendretab, readOnly=True)
            self.rergot.resize(100, 20)
            self.rergot.move(530, 209)
            self.rergot.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            #########################Graines nuisibles (%)##########
            self.rgrainnuisible = QtWidgets.QLineEdit("", self.bltendretab, readOnly=True)
            self.rgrainnuisible.resize(100, 20)
            self.rgrainnuisible.setFont(self.txtpsfont)
            self.rgrainnuisible.move(530, 230)
            self.rgrainnuisible.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            #############################Débris végétaux (%)########
            self.rdébrisvé = QtWidgets.QLineEdit("", self.bltendretab, readOnly=True)
            self.rdébrisvé.resize(100, 20)
            self.rdébrisvé.move(530, 251)
            self.rdébrisvé.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            #########################Matière inerte (%)################
            self.rmatiéreinrt = QtWidgets.QLineEdit("", self.bltendretab, readOnly=True)
            self.rmatiéreinrt.resize(100, 20)
            self.rmatiéreinrt.move(530, 272)
            self.rmatiéreinrt.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            ################################Grains chauffés (%)############################
            self.rgrainchaufé = QtWidgets.QLineEdit("", self.bltendretab, readOnly=True)
            self.rgrainchaufé.resize(100, 20)
            self.rgrainchaufé.move(530, 293)
            self.rgrainchaufé.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            ########################################Grains sans valeur (%)#######################################
            self.rgrainsanvaleur = QtWidgets.QLineEdit("", self.bltendretab, readOnly=True)
            self.rgrainsanvaleur.resize(100, 20)
            self.rgrainsanvaleur.move(530, 314)
            self.rgrainsanvaleur.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            #######################################Grains cariés##########################################
            self.rgraincarré = QtWidgets.QLineEdit("", self.bltendretab, readOnly=True)
            self.rgraincarré.resize(100, 20)
            self.rgraincarré.move(530, 335)
            self.rgraincarré.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            #######################################Total(%) 1er#####################################################
            self.rtotalprem = QtWidgets.QDoubleSpinBox(self.bltendretab, readOnly=True)
            self.rtotalprem.setSpecialValueText(" ")
            self.rtotalprem.resize(100, 20)
            self.rtotalprem.move(530, 356)
            self.rtotalprem.setFont(self.txtpsfont)
            self.rtotalprem.setStyleSheet(" border: 2px solid bleu;border-radius: 4px;padding: 2px")
            ##############################################Grains cassés (%) #########################################################
            self.rgraincassé = QtWidgets.QDoubleSpinBox(self.bltendretab, readOnly=True)
            self.rgraincassé.move(530, 377)
            self.rgraincassé.resize(100, 20)
            self.rgraincassé.setSpecialValueText(' ')
            self.rgraincassé.setFont(self.txtpsfont)
            self.rgraincassé.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            #########################################################Gains échaudés (%)#####################################################
            self.rgrainechaude = QtWidgets.QLineEdit("", self.bltendretab, readOnly=True)
            self.rgrainechaude.move(530, 419)
            self.rgrainechaude.resize(100, 20)
            self.rgrainechaude.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            #####################################################Grains maigres (%)########################################################
            self.rgrainmaigre = QtWidgets.QLineEdit("", self.bltendretab, readOnly=True)
            self.rgrainmaigre.move(530, 398)
            self.rgrainmaigre.resize(100, 20)
            self.rgrainmaigre.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            ##########################################################Grains germés (%)###################################################
            self.rgraingermé = QtWidgets.QLineEdit("", self.bltendretab, readOnly=True)
            self.rgraingermé.move(530, 440)
            self.rgraingermé.resize(100, 20)
            self.rgraingermé.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            ##########################################################Grain punaisés (%)#########################################################
            self.rgrainpunaisé = QtWidgets.QLineEdit("", self.bltendretab, readOnly=True)
            self.rgrainpunaisé.move(530, 461)
            self.rgrainpunaisé.resize(100, 20)
            self.rgrainpunaisé.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            #######################################################################Grains piqués (%)##########################################
            self.rgrainpiqué = QtWidgets.QLineEdit("", self.bltendretab, readOnly=True)
            self.rgrainpiqué.move(530, 482)
            self.rgrainpiqué.resize(100, 20)
            self.rgrainpiqué.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            #######################################################################Grains boutés « faible » (%)#######################################
            self.rgrainboutef = QtWidgets.QLineEdit("", self.bltendretab, readOnly=True)
            self.rgrainboutef.move(530, 503)
            self.rgrainboutef.resize(100, 20)
            self.rgrainboutef.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            ####################################################################Grains boutés  « forte » (%)######################################
            self.rgrainbouté = QtWidgets.QLineEdit("", self.bltendretab, readOnly=True)
            self.rgrainbouté.move(530, 524)
            self.rgrainbouté.resize(100, 20)
            self.rgrainbouté.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            ##################################################Grains mouchetés (%)########################################################
            self.rgrainmouchté = QtWidgets.QLineEdit("", self.bltendretab, readOnly=True)
            self.rgrainmouchté.move(530, 545)
            self.rgrainmouchté.resize(100, 20)
            self.rgrainmouchté.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            ########################################################Grain étrangers Utilisables pour le bétail (%)##################################################
            self.rgrainetrangé = QtWidgets.QLineEdit("", self.bltendretab, readOnly=True)
            self.rgrainetrangé.move(530, 566)
            self.rgrainetrangé.resize(100, 20)
            self.rgrainetrangé.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            ######################################################################Total(%)######################################
            self.rtotaldem = QtWidgets.QDoubleSpinBox(self.bltendretab, readOnly=True)
            self.rtotaldem.move(530, 587)
            self.rtotaldem.setSpecialValueText(" ")
            self.rtotaldem.resize(100, 20)
            self.rtotaldemfont = QtGui.QFont("color:black")
            self.rtotaldemfont.setBold(True)
            self.rtotaldemfont.setPointSize(12)
            self.rtotaldem.setFont(self.rtotaldemfont)
            self.rtotaldem.setStyleSheet(
                "background-color:#ffffff;color:000000;border: 2px solid bleu;border-radius: 4px;padding: 0px")

            self.rtotalboni = QtWidgets.QDoubleSpinBox(self.bltendretab, readOnly=True)
            self.rtotalboni.setSpecialValueText(' ')
            self.rtotalboni.move(530, 608)
            self.rtotalboni.resize(100, 25)
            self.rtotalboni.setFont(self.rtotaldemfont)
            self.rtotalboni.setStyleSheet(
                "background-color:#f67570;color:#000000;border: 2px solid bleu;border-radius: 4px;padding: 0px")
            ###################################################bonifécation#####################################################
            #################label valeure##############
            self.bvaleur = QtWidgets.QLabel("Bonification", self.bltendretab, )
            self.bvaleur.move(710, 145)
            self.bvaleur.resize(80, 20)
            self.bvaleur.setFont(self.font)
            ######################Limites(sans bon ni réf)################
            self.bps = QtWidgets.QDoubleSpinBox(self.bltendretab, readOnly=True)
            self.bps.setFont(self.txtpsfont)
            self.bps.setSpecialValueText(' ')
            self.bps.resize(100, 20)
            self.bps.move(710, 167)
            self.bps.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            ###############################humidite#############
            self.bhumidite = QtWidgets.QDoubleSpinBox(self.bltendretab, readOnly=True)
            self.bhumidite.setSpecialValueText(' ')
            self.bhumidite.resize(100, 20)
            self.bhumidite.move(710, 188)
            self.bhumidite.setFont(self.txtpsfont)
            self.bhumidite.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            #######################ergot#########################
            self.bergot = QtWidgets.QLineEdit("", self.bltendretab, readOnly=True)
            self.bergot.resize(100, 20)
            self.bergot.move(710, 209)
            self.bergot.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            #########################Graines nuisibles (%)##########
            self.bgrainnuisible = QtWidgets.QLineEdit("", self.bltendretab, readOnly=True)
            self.bgrainnuisible.resize(100, 20)
            self.bgrainnuisible.move(710, 230)
            self.bgrainnuisible.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            #############################Débris végétaux (%)########
            self.bdébrisvé = QtWidgets.QLineEdit("", self.bltendretab, readOnly=True)
            self.bdébrisvé.resize(100, 20)
            self.bdébrisvé.move(710, 251)
            self.bdébrisvé.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            #########################Matière inerte (%)################
            self.bmatiéreinrt = QtWidgets.QLineEdit("", self.bltendretab, readOnly=True)
            self.bmatiéreinrt.resize(100, 20)
            self.bmatiéreinrt.move(710, 272)
            self.bmatiéreinrt.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            ################################Grains chauffés (%)############################
            self.bgrainchaufé = QtWidgets.QLineEdit("", self.bltendretab, readOnly=True)
            self.bgrainchaufé.resize(100, 20)
            self.bgrainchaufé.move(710, 293)
            self.bgrainchaufé.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            ########################################Grains sans valeur (%)#######################################
            self.bgrainsanvaleur = QtWidgets.QLineEdit("", self.bltendretab, readOnly=True)
            self.bgrainsanvaleur.resize(100, 20)
            self.bgrainsanvaleur.move(710, 314)
            self.bgrainsanvaleur.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            #######################################Grains cariés##########################################
            self.bgraincarré = QtWidgets.QLineEdit("", self.bltendretab, readOnly=True)
            self.bgraincarré.resize(100, 20)
            self.bgraincarré.move(710, 335)
            self.bgraincarré.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            #######################################Total(%) 1er#####################################################
            self.btotalprem = QtWidgets.QDoubleSpinBox(self.bltendretab, readOnly=True)
            self.btotalprem.setSpecialValueText(' ')
            self.btotalprem.resize(100, 20)
            self.btotalprem.move(710, 356)
            self.btotalprem.setFont(self.txtpsfont)
            self.btotalprem.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            ##############################################Grains cassés (%) #########################################################
            self.bgraincassé = QtWidgets.QLineEdit("", self.bltendretab, readOnly=True)
            self.bgraincassé.move(710, 377)
            self.bgraincassé.resize(100, 20)
            self.graincassé.setFont(self.txtpsfont)
            self.bgraincassé.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            #########################################################Gains échaudés (%)#####################################################
            self.bgrainechaude = QtWidgets.QLineEdit("", self.bltendretab, readOnly=True)
            self.bgrainechaude.move(710, 419)
            self.bgrainechaude.resize(100, 20)
            self.bgrainechaude.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            #####################################################Grains maigres (%)########################################################
            self.bgrainmaigre = QtWidgets.QLineEdit("", self.bltendretab, readOnly=True)
            self.bgrainmaigre.move(710, 398)
            self.bgrainmaigre.resize(100, 20)
            self.bgrainmaigre.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            ##########################################################Grains germés (%)###################################################
            self.bgraigermé = QtWidgets.QLineEdit("", self.bltendretab, readOnly=True)
            self.bgraigermé.move(710, 440)
            self.bgraigermé.resize(100, 20)
            self.bgraigermé.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            ##########################################################Grain punaisés (%)#########################################################
            self.bgrainpunaisé = QtWidgets.QLineEdit("", self.bltendretab, readOnly=True)
            self.bgrainpunaisé.move(710, 461)
            self.bgrainpunaisé.resize(100, 20)
            self.bgrainpunaisé.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            #######################################################################Grains piqués (%)##########################################
            self.bgrainpiqué = QtWidgets.QLineEdit("", self.bltendretab, readOnly=True)
            self.bgrainpiqué.move(710, 482)
            self.bgrainpiqué.resize(100, 20)
            self.bgrainpiqué.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            #######################################################################Grains boutés « faible » (%)#######################################
            self.bgrainboutef = QtWidgets.QLineEdit("", self.bltendretab, readOnly=True)
            self.bgrainboutef.move(710, 503)
            self.bgrainboutef.resize(100, 20)
            self.bgrainboutef.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            ####################################################################Grains boutés  « forte » (%)######################################
            self.bgrainbouté = QtWidgets.QLineEdit("", self.bltendretab, readOnly=True)
            self.bgrainbouté.move(710, 524)
            self.bgrainbouté.resize(100, 20)
            self.bgrainbouté.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            ##################################################Grains mouchetés (%)########################################################
            self.bgrainmouchté = QtWidgets.QLineEdit("", self.bltendretab, readOnly=True)
            self.bgrainmouchté.move(710, 545)
            self.bgrainmouchté.resize(100, 20)
            self.bgrainmouchté.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            ########################################################Grain étrangers Utilisables pour le bétail (%)##################################################
            self.bgrainetrangé = QtWidgets.QLineEdit("", self.bltendretab, readOnly=True)
            self.bgrainetrangé.move(710, 566)
            self.bgrainetrangé.resize(100, 20)
            self.bgrainetrangé.setFont(self.txtpsfont)
            self.bgrainetrangé.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            ######################################################################Total(%)#####################################
            self.btotaldem = QtWidgets.QDoubleSpinBox(self.bltendretab, readOnly=True)
            self.btotaldem.setSpecialValueText(' ')
            self.btotaldem.move(710, 587)
            self.btotaldem.resize(100, 20)
            self.btotaldem.setFont(self.rtotaldemfont)
            self.btotaldem.setStyleSheet(
                "background-color:#ffffff;color:#000000;border: 2px solid bleu;border-radius: 4px;padding: 0px")

            self.btotalboni = QtWidgets.QDoubleSpinBox(self.bltendretab, readOnly=True)
            self.btotalboni.setSpecialValueText(' ')
            self.btotalboni.move(710, 608)
            self.btotalboni.resize(100, 25)
            self.btotalboni.setFont(self.rtotaldemfont)
            self.btotalboni.setStyleSheet(
                "background-color:#9fdcb7;color:#000000;border: 2px solid bleu;border-radius: 4px;padding: 0px")

            #########################################################observation###############
            self.observation = QtWidgets.QLabel("Observation", self.bltendretab)
            self.observation.move(890, 145)
            self.observation.resize(100, 20)
            self.observation.setFont(self.txtpsfont)
            self.observation.setFont(self.font)
            ##################################################txtobservation##################################
            self.txtobservation = QtWidgets.QTextEdit("<h2><h2/>  <h2><h2/>  <h2><h2/> <h2><h2/>   <h3><h3/>",
                                                      self.bltendretab)
            self.txtobservation.move(890, 167)
            self.txtobservation.resize(135, 459)
            self.txtobservation.setStyleSheet("border: 2px solid bleu ;border-radius: 4px;padding: 2px")
            ###################################################label ccls relizane#################
            self.labelccls = QtWidgets.QLabel("<h1>CCLS RELIZANE SERVICE QUALITE<h1/>", self.bltendretab)
            self.labelccls.move(500, 0)
            self.labelccls.resize(438, 90)
            self.labelccls.setFont(self.font)
            self.labelccls.setStyleSheet(
                "background-color: #D8F9DB; border: 2px solid bleu ;border-radius:8px;padding: 0px")
            self.LABELBULLETIN = QtWidgets.QLabel("<H2>BULLETIN MOULIN</H2>", self.bltendretab)
            self.LABELBULLETIN.move(650, 35)
            self.LABELBULLETIN.resize(180, 30)
            self.LABELBULLETIN.setStyleSheet("background-color: #D8F9DB")

            self.bletendretxt = QtWidgets.QLabel("<H2>Blé Tendre<H2/>", self.bltendretab)
            self.bletendretxt.move(690, 60)
            self.bletendretxt.resize(140, 23)
            self.bletendretxt.setStyleSheet("background-color: #D8F9DB")

            #############################################date edit#############################################
            self.dateeditetxt = QtWidgets.QLabel("Date:", self.bltendretab)
            self.dateeditetxt.setGeometry(QtCore.QRect(20, 20, 100, 23))
            self.dateeditetxt.setFont(self.font)
            self.dateday = QDate.currentDate()
            self.dateedite = QtWidgets.QDateEdit(self.bltendretab)
            self.dateedite.setDate(self.dateday)
            self.dateedite.move(84, 20)
            self.dateedite.setStyleSheet(
                " background-color: #ffffff;padding: 1px;border-style: solid;border: 1px solid #76797C;border-color:#000000;border-radius: 0px;color: #000000;")
            self.dateedite.resize(130, 30)
            self.dateedite.setFont(self.font)

            ###############################################search#################################################

            #####################################décade######################
            self.decade = QtWidgets.QLabel("Décade:", self.bltendretab)
            self.decade.move(20, 60)
            self.decade.resize(60, 20)
            self.decade.setFont(self.font)
            self.decadecombo = QtWidgets.QComboBox(self.bltendretab)
            self.decadecombo.addItem("")
            self.decadecombo.addItem("1 ére")
            self.decadecombo.addItem("2 éme")
            self.decadecombo.addItem("3 éme")
            self.decadecombo.move(84, 60)
            self.decadecombo.resize(130, 30)

            ##################################################quantite###############################################
            self.quantite = QtWidgets.QLabel("Quantité(QX):", self.bltendretab)
            self.quantite.move(840, 100)
            self.quantite.resize(85, 40)
            self.quantite.setFont(self.font)
            self.quantitetxt = QtWidgets.QLineEdit("", self.bltendretab)

            self.quantitetxt.move(928, 105)
            self.quantitetxt.resize(80, 30)
            self.quantitetxt.setValidator(QDoubleValidator(0.99, 99.99, 2))
            # self.quantitetxt.setStyleSheet("background-color: #31363b")

            ####################################################éspece###########################
            self.éspéce = QtWidgets.QLabel("Espèce :", self.bltendretab)
            self.éspéce.move(20, 100)
            self.éspéce.resize(120, 40)
            self.éspéce.setFont(self.font)
            self.éspécecombo = QtWidgets.QComboBox(self.bltendretab)
            self.éspécecombo.addItem("Blé Tendre")
            self.éspécecombo.move(84, 105)
            self.éspécecombo.resize(130, 30)
            ##########################################Nom de l’acheteur : moulin###########################################
            self.moulin = QtWidgets.QLabel("Nom de l’acheteur:", self.bltendretab)
            self.moulin.move(220, 100)
            self.moulin.resize(118, 40)
            self.moulin.setFont(self.font)
            self.moulincombo = QtWidgets.QComboBox(self.bltendretab)
            self.moulincombo.addItem("")
            self.moulincombo.addItem("EURL DJERBIR INDUSTRIELE")
            self.moulincombo.addItem("SARL MOULIN O_ABBES")
            self.moulincombo.addItem("EURL MOULIN BELACEL")
            self.moulincombo.addItem("MOULIN TAHAR MESSAOUD")
            self.moulincombo.addItem("SARL MOULIN BENABDELLAH")
            self.moulincombo.addItem("SARL MATAHIN EL HARAMAIN")
            self.moulincombo.addItem("MINOTERIE NOUR EL HAYAT")
            self.moulincombo.addItem("MOULIN MERINE SASSI")
            self.moulincombo.addItem("SARL DJENDLI")
            self.moulincombo.addItem("SARL TRX HYDRO BENHADJAR")
            self.moulincombo.addItem("EURL MOULIN AIN RAHMA")
            self.moulincombo.addItem("MOULIN FARINE BLANCHE")
            self.moulincombo.addItem("MOULIN MAAMAR BENHADJAR")
            self.moulincombo.addItem("MOULIN OULED BENAICHOUCHE")
            self.moulincombo.addItem("SARL FARINIERE DE L’OUEST")
            self.moulincombo.addItem("SARL MATAHINE ADJINE")
            self.moulincombo.addItem("MOULIN CHOUIKH YOUCEF")
            self.moulincombo.addItem("EURL MOULIN DAMAKO")
            self.moulincombo.addItem("SARL MATAHINE SIDI ABDELHADI")
            self.moulincombo.addItem("SARL MATAHINE MINA")
            self.moulincombo.addItem("EURL ELFORSANE PRODUCTION")
            self.moulincombo.addItem("SARL MATAHINE TOUFIK")
            self.moulincombo.move(340, 105)
            self.moulincombo.resize(220, 30)

            #####################################################Point de collecte : #######################################################
            self.pointcollecte = QtWidgets.QLabel("Point de collecte:", self.bltendretab)
            self.pointcollecte.move(570, 100)
            self.pointcollecte.resize(106, 40)
            self.pointcollecte.setFont(self.font)
            self.pointcollectecombo = QtWidgets.QComboBox(self.bltendretab)
            self.pointcollectecombo.addItem("")
            self.pointcollectecombo.addItem("DOCK SILO CENTRAL")
            self.pointcollectecombo.move(680, 105)
            self.pointcollectecombo.resize(150, 30)

            ######################################################Nom de l’Agréeur#######################################################

            self.agréeeur = QtWidgets.QLabel("Nom de l’Agréeur:", self.bltendretab)
            self.agréeeur.move(1015, 100)
            self.agréeeur.resize(112, 40)
            self.agréeeur.setFont(self.font)
            self.agréeeurcombo = QtWidgets.QComboBox(self.bltendretab, editable=False)
            self.agréeeurcombo.addItem("")
            self.agréeeurcombo.addItem("FELOUAH OMAR")
            self.agréeeurcombo.addItem("BEKHEDDA AEK")
            self.agréeeurcombo.addItem("BENAISSA YOUCEF")
            self.agréeeurcombo.addItem("REZZAG SOFIANE ")
            self.agréeeurcombo.addItem("BELBACHA M.NADIR")
            self.agréeeurcombo.move(1133, 105)
            self.agréeeurcombo.resize(147, 30)

            ############################################docx2pdf######################

            self.tamerbt = QTimer()
            self.tamerbt.timeout.connect(self.all_calcul)
            self.tamerbt.setInterval(1000)
            self.tamerbt.start()

            ###########################buttons################

            self.btnsavebt = QtWidgets.QPushButton("ENREGISTRER", self.bltendretab, clicked=lambda: self.docx_file())
            self.btnsavebt.move(1100, 167)
            self.btnsavebt.resize(200, 80)
            self.btnsavebt.setFont(self.font)
            self.btnsavebt.setIcon(QIcon("images/savepis.png"))
            self.btnsavebt.setIconSize(QSize(70, 80))
            # self.btnsavebt.clicked.connect(self.docx_file)

            self.btnprintbt = QtWidgets.QPushButton("IMPRIMER", self.bltendretab, clicked=lambda: self.printer())
            self.btnprintbt.move(1100, 272)
            self.btnprintbt.resize(200, 80)
            self.btnprintbt.setFont(self.font)
            self.btnprintbt.setIcon(QIcon("images/print125.png"))
            self.btnprintbt.setIconSize(QSize(70, 80))
            # btnprint.clicked.connect(printer)

            self.btnefacebt = QtWidgets.QPushButton("EFACER", self.bltendretab, clicked=lambda: self.clear_all())
            self.btnefacebt.move(1100, 377)
            self.btnefacebt.resize(200, 80)
            self.btnefacebt.setIcon(QIcon("images/eraser45877.png"))
            self.btnefacebt.setIconSize(QSize(70, 80))
            self.btnefacebt.setFont(self.font)
            # self.btnefacebt.clicked.connect(self.clear_all)

            # self.btnsearchbt.setFont(self.font)

            # self.btncalculbt = QtWidgets.QPushButton(" CALCULER", self.bltendretab,clicked=lambda :self.all_calcul())
            # self.btncalculbt.move(1120, 525)
            # self.btncalculbt.resize(200, 80)
            # self.btncalculbt.setFont(self.font)
            # self.btncalculbt.setIcon((QIcon("images/calcul12544.png")))
            # self.btncalculbt.setIconSize(QSize(70, 80))
            # self.btncalculbt.clicked.connect(self.all_calcul)

            #############################################BLE DUR
            ########################################################
            #####################################################################
            ###############################################################################
            ##############################################################################################
            self.bldurtab = QtWidgets.QWidget()
            self.bldurtab.setObjectName("bldurtab")
            self.moulinwidget.addTab(self.bldurtab, "")
            self.bldurtab.setStyleSheet("""QToolTip
    {
        border: 1px solid #76797C;
        background-color:  #fff8b0;;
        color: white;
        padding: 5px;
        opacity: 200;
    }

    QWidget
    {
        color: #000000;
        background-color:  #ffaaaa;
        selection-background-color:#3daee9;
        selection-color: #eff0f1;
        background-clip: border;
        border-image: none;
        border: 0px transparent black;
        outline: 0;
    }

    QWidget:item:hover
    {
        background-color: #3daee9;
        color: #eff0f1;
    }

    QWidget:item:selected
    {
        background-color: #3daee9;
    }



    QWidget:disabled
    {
        color: #454545;
        background-color: #31363b;
    }

    QAbstractItemView
    {
        alternate-background-color: #31363b;
        color: #eff0f1;
        border: 1px solid 3A3939;
        border-radius: 2px;
    }

    QWidget:focus, QMenuBar:focus
    {
        border: 1px solid #3daee9;
    }

    QTabWidget:focus, QCheckBox:focus, QRadioButton:focus, QSlider:focus
    {
        border: none;
    }

    QLineEdit
    {
        background-color: #ffffff;
        padding: 1px;
        border-style: solid;
        border: 1px solid #000000;
        border-radius: 2px;
        color: #000000;
        font-size:12px;
        font-weight:bold;
    }
    QDoubleSpinBox
    {
        background-color: #ffffff;
        padding: 1px;
        border-style: solid;
        border: 2px solid #000000;
        border-radius: 4px;
        color:#000000;
        font-size:12px;
        font-weight:bold;

    }
    QDoubleSpinBox:focus
    {
        background-color: #f2f2f2;
        border-style: solid;
        border: 2px solid #76797C;
        border-radius: 4px;
        border-color: #ff8c00;
    }

    QDoubleSpinBox::drop-down
    {
        subcontrol-origin: padding;
        subcontrol-position: top right;
        width: 0px;

        border-left-width: 0px;
        border-left-color: #232629;
        border-left-style: solid;
        border-top-right-radius: 1px;
        border-bottom-right-radius: 1px;
    }



    QGroupBox {
        border:1px solid #76797C;
        border-radius: 2px;
        margin-top: 20px;
    }

    QGroupBox::title {
        subcontrol-origin: margin;
        subcontrol-position: top center;
        padding-left: 10px;
        padding-right: 10px;
        padding-top: 10px;
    }

    QAbstractScrollArea
    {
        border-radius: 2px;
        border: 1px solid #76797C;
        background-color: transparent;
    }

    QScrollBar:horizontal
    {
        height: 15px;
        margin: 3px 15px 3px 15px;
        border: 1px transparent #2A2929;
        border-radius: 4px;
        background-color: #2A2929;
    }

    QScrollBar::handle:horizontal
    {
        background-color: #605F5F;
        min-width: 5px;
        border-radius: 4px;
    }

    QScrollBar::add-line:horizontal
    {
        margin: 0px 3px 0px 3px;
        border-image: url(:/qss_icons/Dark_rc/right_arrow_disabled.png);
        width: 10px;
        height: 10px;
        subcontrol-position: right;
        subcontrol-origin: margin;
    }

    QScrollBar::sub-line:horizontal
    {
        margin: 0px 3px 0px 3px;
        border-image: url(:/qss_icons/Dark_rc/left_arrow_disabled.png);
        height: 10px;
        width: 10px;
        subcontrol-position: left;
        subcontrol-origin: margin;
    }

    QScrollBar::add-line:horizontal:hover,QScrollBar::add-line:horizontal:on
    {
        border-image: url(:/qss_icons/Dark_rc/right_arrow.png);
        height: 10px;
        width: 10px;
        subcontrol-position: right;
        subcontrol-origin: margin;
    }


    QScrollBar::sub-line:horizontal:hover, QScrollBar::sub-line:horizontal:on
    {
        border-image: url(:/qss_icons/Dark_rc/left_arrow.png);
        height: 10px;
        width: 10px;
        subcontrol-position: left;
        subcontrol-origin: margin;
    }

    QScrollBar::up-arrow:horizontal, QScrollBar::down-arrow:horizontal
    {
        background: none;
    }


    QScrollBar::add-page:horizontal, QScrollBar::sub-page:horizontal
    {
        background: none;
    }

    QScrollBar:vertical
    {
        background-color: #2A2929;
        width: 15px;
        margin: 15px 3px 15px 3px;
        border: 1px transparent #2A2929;
        border-radius: 4px;
    }

    QScrollBar::handle:vertical
    {
        background-color: #605F5F;
        min-height: 5px;
        border-radius: 4px;
    }

    QScrollBar::sub-line:vertical
    {
        margin: 3px 0px 3px 0px;
        border-image: url(:/qss_icons/Dark_rc/up_arrow_disabled.png);
        height: 10px;
        width: 10px;
        subcontrol-position: top;
        subcontrol-origin: margin;
    }

    QScrollBar::add-line:vertical
    {
        margin: 3px 0px 3px 0px;
        border-image: url(:/qss_icons/Dark_rc/down_arrow_disabled.png);
        height: 10px;
        width: 10px;
        subcontrol-position: bottom;
        subcontrol-origin: margin;
    }

    QScrollBar::sub-line:vertical:hover,QScrollBar::sub-line:vertical:on
    {

        border-image: url(:/qss_icons/Dark_rc/up_arrow.png);
        height: 10px;
        width: 10px;
        subcontrol-position: top;
        subcontrol-origin: margin;
    }


    QScrollBar::add-line:vertical:hover, QScrollBar::add-line:vertical:on
    {
        border-image: url(:/qss_icons/Dark_rc/down_arrow.png);
        height: 10px;
        width: 10px;
        subcontrol-position: bottom;
        subcontrol-origin: margin;
    }

    QScrollBar::up-arrow:vertical, QScrollBar::down-arrow:vertical
    {
        background: none;
    }


    QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical
    {
        background: none;
    }

    QTextEdit
    {
        background-color: #ffffff;
        color: #000000;
        border: 1px solid #76797C;
        font-size:12px;
        font-weight:bold;
    }

    QPlainTextEdit
    {
        background-color: #232629;;
        color: #eff0f1;
        border-radius: 2px;
        border: 1px solid #76797C;
    }

    QHeaderView::section
    {
        background-color: #76797C;
        color: #eff0f1;
        padding: 5px;
        border: 1px solid #76797C;
    }

    QSizeGrip {
        image: url(:/qss_icons/Dark_rc/sizegrip.png);
        width: 12px;
        height: 12px;
    }


    QMainWindow::separator
    {
        background-color: #31363b;
        color: white;
        padding-left: 4px;
        spacing: 2px;
        border: 1px dashed #76797C;
    }

    QMainWindow::separator:hover
    {

        background-color: #787876;
        color: white;
        padding-left: 4px;
        border: 1px solid #76797C;
        spacing: 2px;
    }


    QMenu::separator
    {
        height: 1px;
        background-color: #76797C;
        color: white;
        padding-left: 4px;
        margin-left: 10px;
        margin-right: 5px;
    }


    QFrame
    {
        border-radius: 2px;
        border: 1px solid #76797C;
    }

    QFrame[frameShape="0"]
    {
        border-radius: 2px;
        border: 1px transparent #76797C;
    }

    QStackedWidget
    {
        border: 1px transparent black;
    }


    QPushButton
    {
        color: #000000;
        background-color:#fcfcfc;
        border-width: 1px;
        border-color: #1e1e1e;
        border-style: solid;
        border-radius: 6;
        padding: 3px;
        font-size: 12px;
        padding-left: 5px;
        padding-right: 5px;
        min-width: 40px;

    }

    QPushButton:disabled
    {
        background-color: #feedee;
        border-width: 1px;
        border-color: #454545;
        border-style: solid;
        padding-top: 5px;
        padding-bottom: 5px;
        padding-left: 10px;
        padding-right: 10px;
        border-radius: 2px;
        color: #454545;
    }

    QPushButton:focus {
        background-color: #feedee;
        color: #000000;
    }

    QPushButton:pressed
    {
        background-color: #3daee9;
        padding-top: -15px;
        padding-bottom: -17px;
    }

    QComboBox
    {   
        background-color:#ffffff;
        color:#000000;
        selection-background-color:#000000;
        border-style: solid;
        border: 1px solid #000000;
        border-radius: 2px;
        min-width: 40px;
        font-size:12px;
        font-weight:bold;
    }

    QPushButton:checked{
        background-color: #76797C;
        border-color: #6A6969;
    }

    QComboBox:hover,QDoubleSpinBox:Hover,QPushButton:hover,QAbstractSpinBox:hover,QLineEdit:hover,QTextEdit:hover,QPlainTextEdit:hover,QAbstractView:hover,QTreeView:hover
    {
        border: 1px solid #ff8c00;
        color: #000000;
    }

    QComboBox:on
    {
        padding-top: 1px;
        padding-left: 1px;
        selection-background-color: #000000;
        color:#000000;
    }

    QComboBox QAbstractItemView
    {
        background-color: #ffffff;
        color:#000000;
        border-radius: 2px;
        border: 1px solid #000000;
        selection-background-color: #000000;
    }

    QComboBox::drop-down
    {
        subcontrol-origin: padding;
        subcontrol-position: top right;
        width: 15px;

        border-left-width: 0px;
        border-left-color: #ff8c00;
        border-left-style: solid;
        border-top-right-radius: 1px;
        border-bottom-right-radius: 1px;
    }


    QLabel
    {
        border: 2px solid black;
        font-size:13px;
        font-weight:bold;
    }

    QTabWidget{
        border: 0px transparent black;
    }

    QTabWidget::pane {
        border: 1px solid #76797C;
        padding: 5px;
        margin: 0px;
    }

    QTabBar
    {
        qproperty-drawBase: 0;
        left: 5px; /* move to the right by 5px */
        border-radius: 3px;
    }

    QTabBar:focus
    {
        border: 0px transparent black;
    }

    QTabBar::close-button  {
        image: url(:/qss_icons/Dark_rc/close.png);
        background: transparent;
    }

    QTabBar::close-button:hover
    {
        image: url(:/qss_icons/Dark_rc/close-hover.png);
        background: transparent;
    }

    QTabBar::close-button:pressed {
        image: url(:/qss_icons/Dark_rc/close-pressed.png);
        background: transparent;
    }

    /* TOP TABS */
    QTabBar::tab:top {
        color: #eff0f1;
        border: 1px solid #76797C;
        border-bottom: 1px transparent black;
        background-color: #31363b;
        padding: 5px;
        min-width: 50px;
        border-top-left-radius: 2px;
        border-top-right-radius: 2px;
    }

    QTabBar::tab:top:!selected
    {
        color: #eff0f1;
        background-color: #54575B;
        border: 1px solid #76797C;
        border-bottom: 1px transparent black;
        border-top-left-radius: 2px;
        border-top-right-radius: 2px;    
    }

    QTabBar::tab:top:!selected:hover {
        background-color: #3daee9;
    }

    /* BOTTOM TABS */
    QTabBar::tab:bottom {
        color: #eff0f1;
        border: 1px solid #76797C;
        border-top: 1px transparent black;
        background-color: #31363b;
        padding: 5px;
        border-bottom-left-radius: 2px;
        border-bottom-right-radius: 2px;
        min-width: 50px;
    }

    QTabBar::tab:bottom:!selected
    {
        color: #eff0f1;
        background-color: #54575B;
        border: 1px solid #76797C;
        border-top: 1px transparent black;
        border-bottom-left-radius: 2px;
        border-bottom-right-radius: 2px;
    }

    QTabBar::tab:bottom:!selected:hover {
        background-color: #3daee9;
    }

    /* LEFT TABS */
    QTabBar::tab:left {
        color: #eff0f1;
        border: 1px solid #76797C;
        border-left: 1px transparent black;
        background-color: #31363b;
        padding: 5px;
        border-top-right-radius: 2px;
        border-bottom-right-radius: 2px;
        min-height: 50px;
    }

    QTabBar::tab:left:!selected
    {
        color: #eff0f1;
        background-color: #54575B;
        border: 1px solid #76797C;
        border-left: 1px transparent black;
        border-top-right-radius: 2px;
        border-bottom-right-radius: 2px;
    }

    QTabBar::tab:left:!selected:hover {
        background-color: #3daee9;
    }


    /* RIGHT TABS */
    QTabBar::tab:right {
        color: #eff0f1;
        border: 1px solid #76797C;
        border-right: 1px transparent black;
        background-color: #31363b;
        padding: 5px;
        border-top-left-radius: 2px;
        border-bottom-left-radius: 2px;
        min-height: 50px;
    }

    QTabBar::tab:right:!selected
    {
        color: #eff0f1;
        background-color: #54575B;
        border: 1px solid #76797C;
        border-right: 1px transparent black;
        border-top-left-radius: 2px;
        border-bottom-left-radius: 2px;
    }

    QTabBar::tab:right:!selected:hover {
        background-color: #3daee9;
    }

    QTabBar QToolButton::right-arrow:enabled {
         image: url(:/qss_icons/Dark_rc/right_arrow.png);
     }

     QTabBar QToolButton::left-arrow:enabled {
         image: url(:/qss_icons/Dark_rc/left_arrow.png);
     }

    QTabBar QToolButton::right-arrow:disabled {
         image: url(:/qss_icons/Dark_rc/right_arrow_disabled.png);
     }

     QTabBar QToolButton::left-arrow:disabled {
         image: url(:/qss_icons/Dark_rc/left_arrow_disabled.png);
     }


    QDockWidget {
        background: #31363b;
        border: 1px solid #403F3F;
        titlebar-close-icon: url(:/qss_icons/Dark_rc/close.png);
        titlebar-normal-icon: url(:/qss_icons/Dark_rc/undock.png);
    }

    QDockWidget::close-button, QDockWidget::float-button {
        border: 1px solid transparent;
        border-radius: 2px;
        background: transparent;
    }

    QDockWidget::close-button:hover, QDockWidget::float-button:hover {
        background: rgba(255, 255, 255, 10);
    }

    QDockWidget::close-button:pressed, QDockWidget::float-button:pressed {
        padding: 1px -1px -1px 1px;
        background: rgba(255, 255, 255, 10);
    }


    QSlider::groove:horizontal {
        border: 1px solid #565a5e;
        height: 4px;
        background: #565a5e;
        margin: 0px;
        border-radius: 2px;
    }

    QSlider::handle:horizontal {
        background: #232629;
        border: 1px solid #565a5e;
        width: 16px;
        height: 16px;
        margin: -8px 0;
        border-radius: 9px;
    }

    QSlider::groove:vertical {
        border: 1px solid #565a5e;
        width: 4px;
        background: #565a5e;
        margin: 0px;
        border-radius: 3px;
    }

    QSlider::handle:vertical {
        background: #232629;
        border: 1px solid #565a5e;
        width: 16px;
        height: 16px;
        margin: 0 -8px;
        border-radius: 9px;
    }

    QToolButton {
        background-color: transparent;
        border: 1px transparent #76797C;
        border-radius: 2px;
        margin: 3px;
        padding: 5px;
    }

    QToolButton[popupMode="1"] { /* only for MenuButtonPopup */
     padding-right: 20px; /* make way for the popup button */
     border: 1px #76797C;
     border-radius: 5px;
    }

    QToolButton[popupMode="2"] { /* only for InstantPopup */
     padding-right: 10px; /* make way for the popup button */
     border: 1px #76797C;
    }


    QToolButton:hover, QToolButton::menu-button:hover {
        background-color: transparent;
        border: 1px solid #3daee9;
        padding: 5px;
    }

    QToolButton:checked, QToolButton:pressed,
            QToolButton::menu-button:pressed {
        background-color: #3daee9;
        border: 1px solid #3daee9;
        padding: 5px;
    }

    /* the subcontrol below is used only in the InstantPopup or DelayedPopup mode */
    QToolButton::menu-indicator {
        background-color:ff8c00;
        top: -7px; left: -2px; /* shift it a bit */
    }

    /* the subcontrols below are used only in the MenuButtonPopup mode */
    QToolButton::menu-button {
        border: 1px transparent #76797C;
        border-top-right-radius: 6px;
        border-bottom-right-radius: 6px;
        /* 16px width + 4px for border = 20px allocated above */
        width: 16px;
        outline: none;
    }

    QToolButton::menu-arrow {
       background-color:ff8c00;
    }

    QToolButton::menu-arrow:open {
        border: 1px solid #76797C;
    }

    QPushButton::menu-indicator  {
        subcontrol-origin: padding;
        subcontrol-position: bottom right;
        left: 8px;
    }

    QTableView
    {
        border: 1px solid #76797C;
        gridline-color: #31363b;
        background-color: #232629;
    }


    QTableView, QHeaderView
    {
        border-radius: 0px;
    }

    QTableView::item:pressed, QListView::item:pressed, QTreeView::item:pressed  {
        background: #3daee9;
        color: #eff0f1;
    }

    QTableView::item:selected:active, QTreeView::item:selected:active, QListView::item:selected:active  {
        background: #3daee9;
        color: #eff0f1;
    }


    QHeaderView
    {
        background-color: #31363b;
        border: 1px transparent;
        border-radius: 0px;
        margin: 0px;
        padding: 0px;

    }

    QHeaderView::section  {
        background-color: #31363b;
        color: #eff0f1;
        padding: 5px;
        border: 1px solid #76797C;
        border-radius: 0px;
        text-align: center;
    }

    QHeaderView::section::vertical::first, QHeaderView::section::vertical::only-one
    {
        border-top: 1px solid #76797C;
    }

    QHeaderView::section::vertical
    {
        border-top: transparent;
    }

    QHeaderView::section::horizontal::first, QHeaderView::section::horizontal::only-one
    {
        border-left: 1px solid #76797C;
    }

    QHeaderView::section::horizontal
    {
        border-left: transparent;
    }


    QHeaderView::section:checked
     {
        color: white;
        background-color: #334e5e;
     }

     /* style the sort indicator */
    QHeaderView::down-arrow {
        image: url(:/qss_icons/Dark_rc/down_arrow.png);
    }

    QHeaderView::up-arrow {
        image: url(:/qss_icons/Dark_rc/up_arrow.png);
    }


    QTableCornerButton::section {
        background-color: #31363b;
        border: 1px transparent #76797C;
        border-radius: 0px;
    }

    QToolBox  {
        padding: 5px;
        border: 1px transparent black;
    }

    QToolBox::tab {
        color: #eff0f1;
        background-color: #31363b;
        border: 1px solid #76797C;
        border-bottom: 1px transparent #31363b;
        border-top-left-radius: 5px;
        border-top-right-radius: 5px;
    }

    QToolBox::tab:selected { /* italicize selected tabs */
        font: italic;
        background-color: #31363b;
        border-color: #3daee9;
     }

    QStatusBar::item {
        border: 0px transparent dark;
     }


    QFrame[height="3"], QFrame[width="3"] {
        background-color: #76797C;
    }




    QDateEdit
    {
        selection-background-color:#31363b;
        border-style: solid;
        border: 1px solid #76797C;
        border-radius: 2px;
        padding: 1px;
        min-width: 75px;
    }

    QDateEdit:on
    {
        padding-top: 2px;
        padding-left: 2px;
        selection-background-color: #4a4a4a;
    }

    QDateEdit QAbstractItemView
    {
        background-color: #ff8c00;
        border-radius: 2px;
        border: 1px solid #3375A3;
        selection-background-color:ff8c00;
    }

    QDateEdit::drop-down
    {
        subcontrol-origin: padding;
        subcontrol-position: top right;
        width: 15px;
        border-left-width: 0px;
        border-left-color: darkgray;
        border-left-style: solid;
        border-top-right-radius: 3px;
        border-bottom-right-radius: 3px;
    }""")
            self.font = QtGui.QFont()
            self.font.setBold(True)
            self.font.setPointSize(10)
            ##########text bul# ettin######
            self.paramétrebd = QtWidgets.QLabel("Paramètre", self.bldurtab)
            self.paramétrebd.move(30, 111)
            self.paramétrebd.resize(80, 20)
            self.paramétrebd.setFont(self.font)
            self.txtpsfontbd = QtGui.QFont()
            self.txtpsfontbd.setBold(True)
            self.txtpsfontbd.setPointSize(9)
            ################Limites(sans bon ni réf)###############
            self.valeurbd = QtWidgets.QLabel("""Limite-ssans-bon-ni-réf)""", self.bldurtab)
            self.valeurbd.move(170, 95)
            self.valeurbd.resize(145, 55)
            self.valeurbd.setFont(self.font)
            ######################Limites(sans bon ni réf)################
            self.psbd = QtWidgets.QLineEdit("Poids spécifique (kg/hl):   (75.500-75.899)", self.bldurtab, readOnly=True)
            self.psbd.resize(319, 20)
            self.psbd.move(30, 131)
            self.psbd.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            self.psbd.setFont(self.txtpsfont)
            self.psbd.setStyleSheet(" border: 2px solid bleu ;border-radius: 4px;padding: 0px")
            ###############################humidite#############
            self.humiditebd = QtWidgets.QLineEdit("Teneur en eau(%):  (13.5-15)", self.bldurtab, readOnly=True)
            self.humiditebd.resize(319, 20)
            self.humiditebd.move(30, 152)
            self.humiditebd.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            self.humiditebd.setFont(self.txtpsfont)

            #######################ergot#########################
            self.ergotbd = QtWidgets.QLineEdit("Ergo(% :<0.001 ", self.bldurtab, readOnly=True)
            self.ergotbd.resize(319, 20)
            self.ergotbd.move(30, 173)
            self.ergotbd.setStyleSheet("background-color: #232629")
            self.ergotbd.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            self.ergotbd.setFont(self.txtpsfont)

            #########################Graines nuisibles (%)##########
            self.grainnuisiblebd = QtWidgets.QLineEdit("Graines nuisibles(%): <0.001", self.bldurtab, readOnly=True)
            self.grainnuisiblebd.resize(319, 20)
            self.grainnuisiblebd.move(30, 194)
            self.grainnuisiblebd.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            self.grainnuisiblebd.setFont(self.txtpsfont)
            #############################Débris végétaux (%)########
            self.débrisvébd = QtWidgets.QLineEdit("Débris végétaux(%):  ", self.bldurtab, readOnly=True)
            self.débrisvébd.resize(319, 20)
            self.débrisvébd.move(30, 215)
            self.débrisvébd.setFont(self.txtpsfont)
            self.débrisvébd.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            #########################Matière inerte (%)################
            self.matiéreinrtbd = QtWidgets.QLineEdit("Matière inerte(%):", self.bldurtab, readOnly=True)
            self.matiéreinrtbd.resize(319, 20)
            self.matiéreinrtbd.move(30, 236)
            self.matiéreinrtbd.setFont(self.txtpsfont)
            self.matiéreinrtbd.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            ################################Grains chauffés (%)############################
            self.grainchaufébd = QtWidgets.QLineEdit("Grains chauffés(%): ", self.bldurtab, readOnly=True)
            self.grainchaufébd.resize(319, 20)
            self.grainchaufébd.move(30, 257)
            self.grainchaufébd.setFont(self.txtpsfont)
            self.grainchaufébd.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            ########################################Grains sans valeur (%)#######################################
            self.grainsanvaleurbd = QtWidgets.QLineEdit("Grains sans valeur(%):", self.bldurtab, readOnly=True)
            self.grainsanvaleurbd.resize(319, 20)
            self.grainsanvaleurbd.move(30, 278)
            self.grainsanvaleurbd.setFont(self.txtpsfont)
            self.grainsanvaleurbd.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            #######################################Grains cariés##########################################
            self.graincarré = QtWidgets.QLineEdit("Grains cariés:   ", self.bldurtab, readOnly=True)
            self.graincarré.resize(319, 20)
            self.graincarré.move(30, 299)
            self.graincarré.setFont(self.txtpsfont)
            self.graincarré.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")

            #######################################Total(%) 1er#####################################################
            self.totalprembd = QtWidgets.QLineEdit("TOTAL 1ére CAT:   ≤1", self.bldurtab, readOnly=True)
            self.totalprembd.resize(319, 20)
            self.totalprembd.move(30, 320)
            self.totalprembd.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            self.totalprembd.setFont(self.txtpsfont)
            ##############################################Grains cassés (%) #########################################################
            self.graincassébd = QtWidgets.QLineEdit("Grains cassés(%):   ≤2", self.bldurtab, readOnly=True)
            self.graincassébd.move(30, 341)
            self.graincassébd.resize(319, 20)
            self.graincassébd.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            self.graincassébd.setFont(self.txtpsfont)
            #########################################################Gains échaudés (%)#####################################################
            self.grainechaudebd = QtWidgets.QLineEdit("Gains échaudés(%):   ", self.bldurtab, readOnly=True)
            self.grainechaudebd.move(30, 362)
            self.grainechaudebd.resize(319, 20)
            self.grainechaudebd.setFont(self.txtpsfont)
            self.grainechaudebd.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            #####################################################Grains maigres (%)########################################################
            self.grainmaigrebd = QtWidgets.QLineEdit("Grains maigres(%):", self.bldurtab, readOnly=True)
            self.grainmaigrebd.move(30, 383)
            self.grainmaigrebd.resize(319, 20)
            self.grainmaigrebd.setFont(self.txtpsfont)
            self.grainmaigrebd.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            ##########################################################Grains germés (%)###################################################
            self.graigermébd = QtWidgets.QLineEdit("Grains germés(%): ≤2", self.bldurtab, readOnly=True)
            self.graigermébd.move(30, 404)
            self.graigermébd.resize(319, 20)
            self.graigermébd.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            self.graigermébd.setFont(self.txtpsfont)
            ##########################################################Grain punaisés (%)#########################################################
            self.grainpunaisébd = QtWidgets.QLineEdit("Grain punaisés(%): ≤1", self.bldurtab, readOnly=True)
            self.grainpunaisébd.move(30, 425)
            self.grainpunaisébd.resize(319, 20)
            self.grainpunaisébd.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            self.grainpunaisébd.setFont(self.txtpsfont)

            #######################################################################Grains piqués (%)##########################################
            self.grainpiquébd = QtWidgets.QLineEdit("Grains piqués(%):  ", self.bldurtab, readOnly=True)
            self.grainpiquébd.move(30, 446)
            self.grainpiquébd.resize(319, 20)
            self.grainpiquébd.setFont(self.txtpsfont)
            self.grainpiquébd.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            #######################################################################Grains boutés « faible » (%)#######################################

            ####################################################################Grains boutés  « forte » (%)######################################
            self.grainboutébd = QtWidgets.QLineEdit("Grains boutés  « forte » (%):", self.bldurtab, readOnly=True)
            self.grainboutébd.move(30, 467)
            self.grainboutébd.resize(319, 20)
            self.grainboutébd.setFont(self.txtpsfont)
            self.grainboutébd.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            ##################################################Grains mouchetés (%)########################################################
            self.grainmouchtébd = QtWidgets.QLineEdit("Grains mouchetés (%):", self.bldurtab, readOnly=True)
            self.grainmouchtébd.move(30, 488)
            self.grainmouchtébd.resize(319, 20)
            self.grainmouchtébd.setFont(self.txtpsfont)
            self.grainmouchtébd.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            ########################################################Grain étrangers Utilisables pour le bétail (%)##################################################
            self.grainetrangébd = QtWidgets.QLineEdit("Grain étrangers Utilisables pour le bétail (%):  ",
                                                      self.bldurtab, readOnly=True)
            self.grainetrangébd.move(30, 509)
            self.grainetrangébd.resize(319, 20)
            self.grainetrangébd.setFont(self.txtpsfont)
            self.grainetrangébd.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            ######################################################################Total(%)######################################
            self.totaldembd = QtWidgets.QLineEdit("Total(%)  Imp2eme cat   ≤5", self.bldurtab, readOnly=True)
            self.totaldembd.move(30, 530)
            self.totaldembd.resize(319, 20)
            self.totaldembd.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            self.totaldembd.setFont(self.txtpsfont)
            #########################################################indice notin##################################################
            self.indicebd = QtWidgets.QLineEdit('Indice Notin ', self.bldurtab, readOnly=True)
            self.indicebd.move(30, 551)
            self.indicebd.resize(319, 20)
            self.indicebd.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            #######################blétendre dans blé dur########
            self.bletendreinbledur = QtWidgets.QLineEdit("Ble tendre dans ble dur(%)", self.bldurtab, readOnly=True)
            self.bletendreinbledur.move(30, 572)
            self.bletendreinbledur.resize(319, 20)
            self.bletendreinbledur.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            ###############################total complet######
            self.totalcomplet = QtWidgets.QLineEdit("TOTAL", self.bldurtab, readOnly=True)
            self.totalcomplet.move(30, 593)
            self.totalcomplet.resize(319, 20)
            self.totalcomplet.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")

            self.totalbonietref = QtWidgets.QLineEdit("\tTotal des Bonifications et Réfactions", self.bldurtab,
                                                      readOnly=True)
            self.totalbonietref.move(30, 614)
            self.totalbonietref.resize(419, 22)
            self.totalbonietref.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            #################label valeure##############
            self.valeurbd = QtWidgets.QLabel("valeur", self.bldurtab)
            self.valeurbd.move(350, 112)
            self.valeurbd.resize(100, 20)
            self.valeurbd.setFont(self.font)
            ######################Limites(sans bon ni réf)################
            self.vpsbd = QtWidgets.QDoubleSpinBox(self.bldurtab)
            self.vpsbd.setRange(71, 84.00)
            self.vpsbd.setSpecialValueText(' ')
            self.vpsbd.resize(100, 20)
            self.vpsbd.move(350, 131)
            self.vpsbd.setFont(self.txtpsfont)
            # self.vpsbd.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            ###############################humidite#############
            self.vhumiditebd = QtWidgets.QDoubleSpinBox(self.bldurtab)
            self.vhumiditebd.setRange(8, 14)
            self.vhumiditebd.resize(100, 20)
            self.vhumiditebd.setSpecialValueText(' ')
            self.vhumiditebd.move(350, 152)
            self.vhumiditebd.setFont(self.txtpsfont)
            # self.vhumiditebd.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            #######################ergot#########################
            self.vergotbd = QtWidgets.QDoubleSpinBox(self.bldurtab, readOnly=False)
            self.vergotbd.setRange(0, 10)
            self.vergotbd.setSpecialValueText(' ')
            self.vergotbd.resize(100, 20)
            self.vergotbd.move(350, 173)
            # self.vergotbd.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            #########################Graines nuisibles (%)##########
            self.vgrainnuisiblebd = QtWidgets.QDoubleSpinBox(self.bldurtab, readOnly=False)
            self.vgrainnuisiblebd.setRange(0, 10)
            self.vgrainnuisiblebd.setSpecialValueText(' ')
            self.vgrainnuisiblebd.resize(100, 20)
            self.vgrainnuisiblebd.move(350, 194)
            # self.vgrainnuisiblebd.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            #############################Débris végétaux (%)########
            self.vdébrisvébd = QtWidgets.QDoubleSpinBox(self.bldurtab, readOnly=False)
            self.vdébrisvébd.setRange(0, 10)
            self.vdébrisvébd.setSpecialValueText(' ')
            self.vdébrisvébd.resize(100, 20)
            self.vdébrisvébd.move(350, 215)
            self.vdébrisvébd.setFont(self.txtpsfont)
            # self.vdébrisvébd.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            #########################Matière inerte (%)################
            self.vmatiéreinrtbd = QtWidgets.QDoubleSpinBox(self.bldurtab, readOnly=False)
            self.vmatiéreinrtbd.setRange(0, 10)
            self.vmatiéreinrtbd.setSpecialValueText(' ')
            self.vmatiéreinrtbd.resize(100, 20)
            self.vmatiéreinrtbd.move(350, 236)
            self.vmatiéreinrtbd.setFont(self.txtpsfont)
            # self.vmatiéreinrtbd.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            ################################Grains chauffés (%)############################
            self.vgrainchaufébd = QtWidgets.QDoubleSpinBox(self.bldurtab, readOnly=False)
            self.vgrainchaufébd.setRange(0, 10)
            self.vgrainchaufébd.setSpecialValueText(' ')
            self.vgrainchaufébd.resize(100, 20)
            self.vgrainchaufébd.move(350, 257)
            self.vgrainchaufébd.setFont(self.txtpsfont)
            # self.vgrainchaufébd.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            ########################################Grains sans valeur (%)#######################################
            self.vgrainsanvaleurbd = QtWidgets.QDoubleSpinBox(self.bldurtab, readOnly=False)
            self.vgrainsanvaleurbd.setSpecialValueText(' ')
            self.vgrainsanvaleurbd.setRange(0, 10)
            self.vgrainsanvaleurbd.resize(100, 20)
            self.vgrainsanvaleurbd.move(350, 278)
            self.vgrainsanvaleurbd.setFont(self.txtpsfont)
            # self.vgrainsanvaleurbd.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            #######################################Grains cariés##########################################
            self.vgraincarrébd = QtWidgets.QDoubleSpinBox(self.bldurtab, readOnly=False)
            self.vgraincarrébd.setSpecialValueText(' ')
            self.vgraincarrébd.setRange(0, 10)
            self.vgraincarrébd.resize(100, 20)
            self.vgraincarrébd.move(350, 299)
            # self.vgraincarrébd.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            #######################################Total(%) 1er#####################################################
            self.vtotalprembd = QtWidgets.QDoubleSpinBox(self.bldurtab, readOnly=True)
            self.vtotalprembd.setSpecialValueText(' ')
            self.vtotalprembd.setRange(0, 10)
            self.vtotalprembd.resize(100, 20)
            self.vtotalprembd.move(350, 320)
            self.vtotalprembd.setFont(self.txtpsfont)
            # self.vtotalprembd.setStyleSheet("background-color:#ffffff;color:#000000;border: 2px solid bleu;border-radius: 4px;padding: 0px")
            ##############################################Grains cassés (%) #########################################################
            self.vgraincassébd = QtWidgets.QDoubleSpinBox(self.bldurtab)
            self.vgraincassébd.move(350, 341)
            self.vgraincassébd.resize(100, 20)
            self.vgraincassébd.setRange(0, 10)
            self.vgraincassébd.setSpecialValueText(" ")
            self.vgraincassébd.setFont(self.txtpsfont)
            # self.vgraincassébd.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 0px")
            #########################################################Gains échaudés (%)#####################################################
            self.vgrainechaudebd = QtWidgets.QDoubleSpinBox(self.bldurtab)
            self.vgrainechaudebd.setSpecialValueText(" ")
            self.vgrainechaudebd.setRange(0, 10)
            self.vgrainechaudebd.move(350, 362)
            self.vgrainechaudebd.resize(100, 20)
            self.vgrainechaudebd.setFont(self.txtpsfont)
            # self.vgrainechaudebd.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            #####################################################Grains maigres (%)########################################################
            self.vgrainmaigrebd = QtWidgets.QDoubleSpinBox(self.bldurtab)
            self.vgrainmaigrebd.setRange(0, 10)
            self.vgrainmaigrebd.setSpecialValueText(" ")
            self.vgrainmaigrebd.move(350, 383)
            self.vgrainmaigrebd.setFont(self.txtpsfont)
            self.vgrainmaigrebd.resize(100, 20)
            # self.vgrainmaigrebd.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            ##########################################################Grains germés (%)###################################################
            self.vgraigermébd = QtWidgets.QDoubleSpinBox(self.bldurtab, readOnly=False)
            self.vgraigermébd.move(350, 404)
            self.vgraigermébd.resize(100, 20)
            self.vgraigermébd.setSpecialValueText('  ')
            self.vgraigermébd.setFont(self.txtpsfont)
            # self.vgraigermébd.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            ##########################################################Grain punaisés (%)#########################################################
            self.vgrainpunaisébd = QtWidgets.QDoubleSpinBox(self.bldurtab, readOnly=False)
            self.vgrainpunaisébd.move(350, 425)
            self.vgrainpunaisébd.resize(100, 20)
            self.vgrainpunaisébd.setSpecialValueText('   ')
            self.vgrainpunaisébd.setFont(self.txtpsfont)
            # self.vgrainpunaisébd.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            #######################################################################Grains piqués (%)##########################################
            self.vgrainpiquébd = QtWidgets.QDoubleSpinBox(self.bldurtab, readOnly=False)
            self.vgrainpiquébd.move(350, 446)
            self.vgrainpiquébd.resize(100, 20)
            self.vgrainpiquébd.setSpecialValueText('  ')
            self.vgrainpiquébd.setFont(self.txtpsfont)
            # self.vgrainpiquébd.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            #######################################################################Grains boutés « faible » (%)#######################################

            ####################################################################Grains boutés  « forte » (%)######################################
            self.vgrainboutébd = QtWidgets.QDoubleSpinBox(self.bldurtab, readOnly=False)
            self.vgrainboutébd.move(350, 467)
            self.vgrainboutébd.resize(100, 20)
            self.vgrainboutébd.setSpecialValueText('  ')
            self.vgrainboutébd.setFont(self.txtpsfont)
            # .vgrainboutébd.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            ##################################################Grains mouchetés (%)########################################################
            self.vgrainmouchtébd = QtWidgets.QDoubleSpinBox(self.bldurtab)
            self.vgrainmouchtébd.move(350, 488)
            self.vgrainmouchtébd.resize(100, 20)
            self.vgrainmouchtébd.setRange(0, 10)
            self.vgrainmouchtébd.setSpecialValueText(' ')
            self.vgrainmouchtébd.setFont(self.txtpsfont)
            # self.vgrainmouchtébd.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            ########################################################Grain étrangers Utilisables pour le bétail (%)##################################################
            self.vgrainetrangébd = QtWidgets.QDoubleSpinBox(self.bldurtab)
            self.vgrainetrangébd.move(350, 509)
            self.vgrainetrangébd.resize(100, 20)
            self.vgrainetrangébd.setRange(0, 10)
            self.vgrainetrangébd.setSpecialValueText(' ')
            self.vgrainetrangébd.setFont(self.txtpsfont)
            # self.vgrainetrangébd.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            ######################################################################Total(%)######################################
            self.vtotaldembd = QtWidgets.QDoubleSpinBox(self.bldurtab, readOnly=True)
            self.vtotaldembd.setRange(1, 20)
            self.vtotaldembd.move(350, 530)
            self.vtotaldembd.resize(100, 20)
            self.vtotaldembd.setSpecialValueText(' ')
            self.vtotaldembd.setFont(self.txtpsfont)
            # self.vtotaldembd.setStyleSheet("background-color:#ffffff;color:#000000;border: 2px solid bleu;border-radius: 4px;padding: 0px")
            ######################################indicenotin #################
            self.vindicenotin = QtWidgets.QDoubleSpinBox(self.bldurtab, readOnly=False)
            self.vindicenotin.setSpecialValueText(' ')
            self.vindicenotin.setFont(self.rtotaldemfont)
            self.vindicenotin.move(350, 551)
            self.vindicenotin.resize(100, 20)
            # self.vindicenotin.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            ########################ble tendre dand blé dur############
            self.vblétendreinbledur = QtWidgets.QDoubleSpinBox(self.bldurtab, readOnly=False)
            self.vblétendreinbledur.setSpecialValueText(" ")
            self.vblétendreinbledur.move(350, 572)
            self.vblétendreinbledur.resize(100, 20)
            # self.vblétendreinbledur.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            #######################bonification total complet##############
            self.vtotalcomplet = QtWidgets.QDoubleSpinBox(self.bldurtab, readOnly=True)
            self.vtotalcomplet.setSpecialValueText(' ')
            self.vtotalcomplet.move(350, 593)
            self.vtotalcomplet.resize(100, 20)
            # self.vtotalcomplet.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            #######################################################réfaction##############################################
            #################label valeure##############
            self.rvaleurbd = QtWidgets.QLabel("Réfaction", self.bldurtab)
            self.rvaleurbd.move(530, 111)
            self.rvaleurbd.resize(80, 20)
            self.rvaleurbd.setFont(self.font)
            ######################Limites(sans bon ni réf)################
            self.rpsbd = QtWidgets.QDoubleSpinBox(self.bldurtab, readOnly=True)
            self.rpsbd.setSpecialValueText(" ")
            self.rpsbd.resize(100, 20)
            self.rpsbd.move(530, 131)
            self.rpsbd.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            ###############################humidite#############
            self.rhumiditebd = QtWidgets.QLineEdit("", self.bldurtab, readOnly=True)
            self.rhumiditebd.resize(100, 20)
            self.rhumiditebd.move(530, 152)
            self.rhumiditebd.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            #######################ergot#########################
            self.rergotbd = QtWidgets.QLineEdit("", self.bldurtab, readOnly=True)
            self.rergotbd.resize(100, 20)
            self.rergotbd.move(530, 173)
            self.rergotbd.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            #########################Graines nuisibles (%)##########
            self.rgrainnuisiblebd = QtWidgets.QLineEdit("", self.bldurtab, readOnly=True)
            self.rgrainnuisiblebd.resize(100, 20)
            self.rgrainnuisiblebd.move(530, 194)
            self.rgrainnuisiblebd.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            #############################Débris végétaux (%)########
            self.rdébrisvébd = QtWidgets.QLineEdit("", self.bldurtab, readOnly=True)
            self.rdébrisvébd.resize(100, 20)
            self.rdébrisvébd.move(530, 215)
            self.rdébrisvébd.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            #########################Matière inerte (%)################
            self.rmatiéreinrtbd = QtWidgets.QLineEdit("", self.bldurtab, readOnly=True)
            self.rmatiéreinrtbd.resize(100, 20)
            self.rmatiéreinrtbd.move(530, 236)
            self.rmatiéreinrtbd.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            ################################Grains chauffés (%)############################
            self.rgrainchaufébd = QtWidgets.QLineEdit("", self.bldurtab, readOnly=True)
            self.rgrainchaufébd.resize(100, 20)
            self.rgrainchaufébd.move(530, 257)
            self.rgrainchaufébd.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            ########################################Grains sans valeur (%)#######################################
            self.rgrainsanvaleurbd = QtWidgets.QLineEdit("", self.bldurtab, readOnly=True)
            self.rgrainsanvaleurbd.resize(100, 20)
            self.rgrainsanvaleurbd.move(530, 278)
            self.rgrainsanvaleurbd.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            #######################################Grains cariés##########################################
            self.rgraincarrébd = QtWidgets.QLineEdit("", self.bldurtab, readOnly=True)
            self.rgraincarrébd.resize(100, 20)
            self.rgraincarrébd.move(530, 299)
            self.rgraincarrébd.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            #######################################Total(%) 1er#####################################################
            self.rtotalprembd = QtWidgets.QDoubleSpinBox(self.bldurtab, readOnly=True)
            self.rtotalprembd.setSpecialValueText(" ")
            self.rtotalprembd.resize(100, 20)
            self.rtotalprembd.move(530, 320)
            self.rtotalprembd.setStyleSheet(" border: 2px solid bleu;border-radius: 4px;padding: 2px")
            ##############################################Grains cassés (%) #########################################################
            self.rgraincassébd = QtWidgets.QDoubleSpinBox(self.bldurtab, readOnly=True)
            self.rgraincassébd.move(530, 341)
            self.rgraincassébd.resize(100, 20)
            self.rgraincassébd.setSpecialValueText(' ')
            self.rgraincassébd.setFont(self.txtpsfont)
            self.rgraincassébd.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            #########################################################Gains échaudés (%)#####################################################
            self.rgrainechaudebd = QtWidgets.QLineEdit("", self.bldurtab, readOnly=True)
            self.rgrainechaudebd.move(530, 362)
            self.rgrainechaudebd.resize(100, 20)
            self.rgrainechaudebd.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            #####################################################Grains maigres (%)########################################################
            self.rgrainmaigrebd = QtWidgets.QLineEdit("", self.bldurtab, readOnly=True)
            self.rgrainmaigrebd.move(530, 383)
            self.rgrainmaigrebd.resize(100, 20)
            self.rgrainmaigrebd.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            ##########################################################Grains germés (%)###################################################
            self.rgraingermébd = QtWidgets.QLineEdit("", self.bldurtab, readOnly=True)
            self.rgraingermébd.move(530, 404)
            self.rgraingermébd.resize(100, 20)
            self.rgraingermébd.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            ##########################################################Grain punaisés (%)#########################################################
            self.rgrainpunaisébd = QtWidgets.QLineEdit("", self.bldurtab, readOnly=True)
            self.rgrainpunaisébd.move(530, 425)
            self.rgrainpunaisébd.resize(100, 20)
            self.rgrainpunaisébd.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            #######################################################################Grains piqués (%)##########################################
            self.rgrainpiquébd = QtWidgets.QLineEdit("", self.bldurtab, readOnly=True)
            self.rgrainpiquébd.move(530, 446)
            self.rgrainpiquébd.resize(100, 20)
            self.rgrainpiquébd.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            #######################################################################Grains boutés « faible » (%)#######################################

            ####################################################################Grains boutés  « forte » (%)######################################
            self.rgrainboutébd = QtWidgets.QDoubleSpinBox(self.bldurtab, readOnly=True)
            self.rgrainboutébd.setSpecialValueText(' ')
            self.rgrainboutébd.move(530, 467)
            self.rgrainboutébd.resize(100, 20)
            self.rgrainboutébd.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            ##################################################Grains mouchetés (%)########################################################
            self.rgrainmouchtébd = QtWidgets.QLineEdit("", self.bldurtab, readOnly=True)
            self.rgrainmouchtébd.move(530, 488)
            self.rgrainmouchtébd.resize(100, 20)
            self.rgrainmouchtébd.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            ########################################################Grain étrangers Utilisables pour le bétail (%)##################################################
            self.rgrainetrangébd = QtWidgets.QLineEdit("", self.bldurtab, readOnly=True)
            self.rgrainetrangébd.move(530, 509)
            self.rgrainetrangébd.resize(100, 20)
            self.rgrainetrangébd.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            ######################################################################Total(%)######################################
            self.rtotaldembd = QtWidgets.QDoubleSpinBox(self.bldurtab, readOnly=True)
            self.rtotaldembd.move(530, 530)
            # self.rtotaldembd.setDecimals(2)
            self.rtotaldembd.setRange(0, 30)
            self.rtotaldembd.setSpecialValueText(" ")
            self.rtotaldembd.resize(100, 20)
            self.rtotaldembd.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            self.rtotaldemfontbd = QtGui.QFont("color:black")
            self.rtotaldemfontbd.setBold(True)
            self.rtotaldemfontbd.setPointSize(12)
            self.rtotaldembd.setFont(self.rtotaldemfont)

            ################refaction indicenotin valeur ###########
            self.rindicenotin = QtWidgets.QDoubleSpinBox(self.bldurtab, readOnly=True)
            self.rindicenotin.setSpecialValueText(' ')
            self.rindicenotin.setFont(self.rtotaldemfont)
            self.rindicenotin.move(530, 551)
            self.rindicenotin.resize(100, 20)
            self.rindicenotin.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            ########################ble tendre dand blé dur############
            self.rblétendreinbledur = QtWidgets.QDoubleSpinBox(self.bldurtab, readOnly=True)
            self.rblétendreinbledur.setSpecialValueText(" ")
            self.rblétendreinbledur.move(530, 572)
            self.rblétendreinbledur.resize(100, 20)
            self.rblétendreinbledur.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            #######################bonification total complet##############
            self.rtotalcomplet = QtWidgets.QDoubleSpinBox(self.bldurtab, readOnly=True)
            self.rtotalcomplet.setSpecialValueText(' ')
            self.rtotalcomplet.move(530, 593)
            self.rtotalcomplet.resize(100, 20)
            self.rtotalcomplet.setStyleSheet(
                "background-color:#ffffff;color:000000;border: 2px solid bleu;border-radius: 4px;padding: 0px")
            self.rtotalcomplet.setFont(self.rtotaldemfont)

            self.rtotalbonibd = QtWidgets.QDoubleSpinBox(self.bldurtab, readOnly=True)
            self.rtotalbonibd.setSpecialValueText(' ')
            self.rtotalbonibd.move(530, 614)
            self.rtotalbonibd.resize(100, 22)
            self.rtotalbonibd.setStyleSheet(
                "background-color:#f67570;color:#000000;border: 2px solid bleu;border-radius: 4px;padding: 0px")
            self.rtotalbonibd.setFont(self.rtotaldemfont)
            ###################################################bonifécation#####################################################
            #################label valeure##############
            self.bvaleurbd = QtWidgets.QLabel("Bonification", self.bldurtab, )
            self.bvaleurbd.move(710, 111)
            self.bvaleurbd.resize(80, 20)
            self.bvaleurbd.setFont(self.font)
            ######################Limites(sans bon ni réf)################
            self.bpsbd = QtWidgets.QDoubleSpinBox(self.bldurtab, readOnly=True)
            self.bpsbd.setFont(self.txtpsfont)
            self.bpsbd.setSpecialValueText(' ')
            self.bpsbd.resize(100, 20)
            self.bpsbd.move(710, 131)
            self.bpsbd.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            ###############################humidite#############
            self.bhumiditebd = QtWidgets.QDoubleSpinBox(self.bldurtab, readOnly=True)
            self.bhumiditebd.setSpecialValueText(' ')
            self.bhumiditebd.resize(100, 20)
            self.bhumiditebd.move(710, 152)
            self.bhumiditebd.setFont(self.txtpsfont)
            self.bhumiditebd.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            #######################ergot#########################
            self.bergotbd = QtWidgets.QLineEdit("", self.bldurtab, readOnly=True)
            self.bergotbd.resize(100, 20)
            self.bergotbd.move(710, 173)
            self.bergotbd.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            #########################Graines nuisibles (%)##########
            self.bgrainnuisiblebd = QtWidgets.QLineEdit("", self.bldurtab, readOnly=True)
            self.bgrainnuisiblebd.resize(100, 20)
            self.bgrainnuisiblebd.move(710, 194)
            self.bgrainnuisiblebd.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            #############################Débris végétaux (%)########
            self.bdébrisvébd = QtWidgets.QLineEdit("", self.bldurtab, readOnly=True)
            self.bdébrisvébd.resize(100, 20)
            self.bdébrisvébd.move(710, 215)
            self.bdébrisvébd.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            #########################Matière inerte (%)################
            self.bmatiéreinrtbd = QtWidgets.QLineEdit("", self.bldurtab, readOnly=True)
            self.bmatiéreinrtbd.resize(100, 20)
            self.bmatiéreinrtbd.move(710, 236)
            self.bmatiéreinrtbd.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            ################################Grains chauffés (%)############################
            self.bgrainchaufé = QtWidgets.QLineEdit("", self.bldurtab, readOnly=True)
            self.bgrainchaufé.resize(100, 20)
            self.bgrainchaufé.move(710, 257)
            self.bgrainchaufé.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            ########################################Grains sans valeur (%)#######################################
            self.bgrainsanvaleurbd = QtWidgets.QLineEdit("", self.bldurtab, readOnly=True)
            self.bgrainsanvaleurbd.resize(100, 20)
            self.bgrainsanvaleurbd.move(710, 278)
            self.bgrainsanvaleurbd.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            #######################################Grains cariés##########################################
            self.bgraincarrébd = QtWidgets.QLineEdit("", self.bldurtab, readOnly=True)
            self.bgraincarrébd.resize(100, 20)
            self.bgraincarrébd.move(710, 299)
            self.bgraincarrébd.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            #######################################Total(%) 1er#####################################################
            self.btotalprembd = QtWidgets.QDoubleSpinBox(self.bldurtab, readOnly=True)
            self.btotalprembd.setSpecialValueText(' ')
            self.btotalprembd.resize(100, 20)
            self.btotalprembd.move(710, 320)
            self.btotalprembd.setFont(self.txtpsfont)
            self.btotalprembd.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            ##############################################Grains cassés (%) #########################################################
            self.bgraincassébd = QtWidgets.QLineEdit("", self.bldurtab, readOnly=True)
            self.bgraincassébd.move(710, 341)
            self.bgraincassébd.resize(100, 20)
            self.bgraincassébd.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            #########################################################Gains échaudés (%)#####################################################
            self.bgrainechaudebd = QtWidgets.QLineEdit("", self.bldurtab, readOnly=True)
            self.bgrainechaudebd.move(710, 362)
            self.bgrainechaudebd.resize(100, 20)
            self.bgrainechaudebd.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            #####################################################Grains maigres (%)########################################################
            self.bgrainmaigrebd = QtWidgets.QLineEdit("", self.bldurtab, readOnly=True)
            self.bgrainmaigrebd.move(710, 383)
            self.bgrainmaigrebd.resize(100, 20)
            self.bgrainmaigrebd.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            ##########################################################Grains germés (%)###################################################
            self.bgraigermébd = QtWidgets.QLineEdit("", self.bldurtab, readOnly=True)
            self.bgraigermébd.move(710, 404)
            self.bgraigermébd.resize(100, 20)
            self.bgraigermébd.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            ##########################################################Grain punaisés (%)#########################################################
            self.bgrainpunaisébd = QtWidgets.QLineEdit("", self.bldurtab, readOnly=True)
            self.bgrainpunaisébd.move(710, 425)
            self.bgrainpunaisébd.resize(100, 20)
            self.bgrainpunaisébd.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            #######################################################################Grains piqués (%)##########################################
            self.bgrainpiquébd = QtWidgets.QLineEdit("", self.bldurtab, readOnly=True)
            self.bgrainpiquébd.move(710, 446)
            self.bgrainpiquébd.resize(100, 20)
            self.bgrainpiquébd.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            #######################################################################Grains boutés « faible » (%)#######################################

            ####################################################################Grains boutés  « forte » (%)######################################
            self.bgrainboutébd = QtWidgets.QLineEdit("", self.bldurtab, readOnly=True)
            self.bgrainboutébd.move(710, 467)
            self.bgrainboutébd.resize(100, 20)
            self.bgrainboutébd.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            ##################################################Grains mouchetés (%)########################################################
            self.bgrainmouchtébd = QtWidgets.QLineEdit("", self.bldurtab, readOnly=True)
            self.bgrainmouchtébd.move(710, 488)
            self.bgrainmouchtébd.resize(100, 20)
            self.bgrainmouchtébd.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            ########################################################Grain étrangers Utilisables pour le bétail (%)##################################################
            self.bgrainetrangébd = QtWidgets.QLineEdit("", self.bldurtab, readOnly=True)
            self.bgrainetrangébd.move(710, 509)
            self.bgrainetrangébd.resize(100, 20)
            self.bgrainetrangébd.setFont(self.txtpsfont)
            self.bgrainetrangébd.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            ######################################################################Total(%)#####################################
            self.btotaldembd = QtWidgets.QDoubleSpinBox(self.bldurtab, readOnly=True)
            self.btotaldembd.setSpecialValueText(' ')
            self.btotaldembd.move(710, 530)
            self.btotaldembd.resize(100, 20)
            self.btotaldembd.setFont(self.rtotaldemfont)
            self.btotaldembd.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")

            ################indicenotin valeur ###########
            self.bindicenotin = QtWidgets.QDoubleSpinBox(self.bldurtab, readOnly=True)
            self.bindicenotin.setSpecialValueText(' ')
            self.bindicenotin.setFont(self.rtotaldemfont)
            self.bindicenotin.move(710, 551)
            self.bindicenotin.resize(100, 20)
            self.bindicenotin.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            ########################ble tendre dand blé dur############
            self.bblétendreinbledur = QtWidgets.QDoubleSpinBox(self.bldurtab, readOnly=True)
            self.bblétendreinbledur.setSpecialValueText(" ")
            self.bblétendreinbledur.setDecimals(3)
            self.bblétendreinbledur.move(710, 572)
            self.bblétendreinbledur.resize(100, 20)
            self.bblétendreinbledur.setStyleSheet("border: 2px solid bleu;border-radius: 4px;padding: 2px")
            #######################bonification total complet##############
            self.btotalcomplet = QtWidgets.QDoubleSpinBox(self.bldurtab, readOnly=True)
            self.btotalcomplet.setSpecialValueText(' ')
            self.btotalcomplet.move(710, 593)
            self.btotalcomplet.resize(100, 20)
            self.btotalcomplet.setStyleSheet(
                "background-color:#ffffff;color:#000000;border: 2px solid bleu;border-radius: 4px;padding: 0px")
            self.btotalcomplet.setFont(self.rtotaldemfont)

            self.btotalbonibd = QtWidgets.QDoubleSpinBox(self.bldurtab, readOnly=True)
            self.btotalbonibd.setSpecialValueText(' ')
            self.btotalbonibd.move(710, 614)
            self.btotalbonibd.resize(100, 22)
            self.btotalbonibd.setStyleSheet(
                "background-color:#9fdcb7;color:#000000;border: 2px solid bleu;border-radius: 4px;padding: 0px")
            self.btotalcomplet.setFont(self.rtotaldemfont)
            #########################################################observation###############
            self.observationbd = QtWidgets.QLabel("Observation", self.bldurtab)
            self.observationbd.move(890, 111)
            self.observationbd.resize(100, 20)
            self.observationbd.setFont(self.txtpsfont)
            self.observationbd.setFont(self.font)
            ##################################################txtobservation##################################
            self.txtobservationbd = QtWidgets.QTextEdit("<h2><h2/>  <h2><h2/>  <h2><h2/> <h2><h2/>   <h3><h3/>",
                                                        self.bldurtab)
            self.txtobservationbd.move(890, 131)
            self.txtobservationbd.resize(135, 503)
            self.txtobservationbd.setStyleSheet("border: 2px solid bleu ;border-radius: 4px;padding: 2px")
            ###################################################label ccls relizane#################
            self.labelcclsbd = QtWidgets.QLabel("<h1>CCLS RELIZANE SERVICE QUALITE<h1/>", self.bldurtab)
            self.labelcclsbd.move(500, 0)
            self.labelcclsbd.resize(438, 80)
            self.labelcclsbd.setFont(self.font)
            self.labelcclsbd.setStyleSheet(
                "background-color: #ffaaaa; border: 2px solid bleu ;border-radius: 6px;padding: 0px")
            self.LABELBULLETINbd = QtWidgets.QLabel("<H2>BULLETIN MOULIN</H2>", self.bldurtab)
            self.LABELBULLETINbd.move(650, 30)
            self.LABELBULLETINbd.resize(180, 23)
            self.LABELBULLETINbd.setStyleSheet("background-color: #ffaaaa")
            self.bledurtxt = QtWidgets.QLabel("<H2>Blé DUR<H2/>", self.bldurtab)
            self.bledurtxt.move(698, 53)
            self.bledurtxt.resize(120, 23)
            self.bledurtxt.setStyleSheet("background-color: #ffaaaa")

            #############################################date edit#############################################
            self.dateeditetxtbd = QtWidgets.QLabel("Date:", self.bldurtab)
            self.dateeditetxtbd.setGeometry(QtCore.QRect(30, 5, 100, 20))
            self.dateeditetxtbd.setFont(self.font)
            self.dateeditebd = QtWidgets.QDateEdit(self.bldurtab)
            self.dateeditebd.setDate(self.dateday)
            self.dateeditebd.move(100, 5)
            self.dateeditebd.resize(112, 25)
            self.dateeditebd.setStyleSheet(
                " background-color:#ffffff;padding: 1px;border-style: solid;border: 1px solid #76797C;border-color:#000000;border-radius: 0px;color: #000000;")
            self.dateeditebd.setFont(self.font)

            ###############################################search#################################################

            #####################################décade######################
            self.decadebd = QtWidgets.QLabel("Décade:", self.bldurtab)
            self.decadebd.move(30, 45)
            self.decadebd.resize(60, 20)
            self.decadebd.setFont(self.font)
            self.decadecombobd = QtWidgets.QComboBox(self.bldurtab)
            self.decadecombobd.addItem("")
            self.decadecombobd.addItem("1 ére")
            self.decadecombobd.addItem("2 éme")
            self.decadecombobd.addItem("3 éme")
            self.decadecombobd.move(100, 45)
            self.decadecombobd.resize(112, 25)

            ##################################################quantite###############################################
            self.quantitebd = QtWidgets.QLabel("Quantité(QX):", self.bldurtab)
            self.quantitebd.move(840, 88)
            self.quantitebd.resize(85, 25)
            self.quantitebd.setFont(self.font)
            self.quantitetxtbd = QtWidgets.QLineEdit("", self.bldurtab)
            self.quantitetxtbd.move(928, 88)
            self.quantitetxtbd.resize(80, 25)
            self.quantitetxtbd.setValidator(QDoubleValidator(0.99, 99.99, 2))

            ####################################################éspece###########################
            self.éspécebd = QtWidgets.QLabel("Espèce :", self.bldurtab)
            self.éspécebd.move(30, 88)
            self.éspécebd.resize(60, 20)
            self.éspécebd.setFont(self.font)
            self.éspécecombobd = QtWidgets.QComboBox(self.bldurtab)
            self.éspécecombobd.addItem("Blé Dur")
            self.éspécecombobd.move(100, 88)
            self.éspécecombobd.resize(112, 25)

            ##########################################Nom de l’acheteur : moulin###########################################
            self.moulinbd = QtWidgets.QLabel("Nom de l’acheteur:", self.bldurtab)
            self.moulinbd.move(220, 88)
            self.moulinbd.resize(118, 20)
            self.moulinbd.setFont(self.font)
            self.moulincombobd = QtWidgets.QComboBox(self.bldurtab)
            self.moulincombobd.addItem("")
            self.moulincombobd.addItem("SARL MOULIN BENABDELLAH")
            self.moulincombobd.move(340, 88)
            self.moulincombobd.resize(220, 25)

            #####################################################Point de collecte : #######################################################
            self.pointcollectebd = QtWidgets.QLabel("Point de collecte:", self.bldurtab)
            self.pointcollectebd.move(570, 80)
            self.pointcollectebd.resize(106, 40)
            self.pointcollectebd.setFont(self.font)
            self.pointcollectecombobd = QtWidgets.QComboBox(self.bldurtab)
            self.pointcollectecombobd.addItem("")
            self.pointcollectecombobd.addItem("DOCK SILO CENTRAL")
            self.pointcollectecombobd.move(680, 88)
            self.pointcollectecombobd.resize(150, 25)

            ######################################################Nom de l’Agréeur#######################################################

            self.agréeeurbd = QtWidgets.QLabel("Nom de l’Agréeur:", self.bldurtab)
            self.agréeeurbd.move(1015, 80)
            self.agréeeurbd.resize(112, 40)
            self.agréeeurbd.setFont(self.font)
            self.agréeeurcombobd = QtWidgets.QComboBox(self.bldurtab, editable=False)
            self.agréeeurcombobd.addItem("")
            self.agréeeurcombobd.addItem("FELOUAH OMAR")
            self.agréeeurcombobd.addItem("BEKHEDDA AEK")
            self.agréeeurcombobd.addItem("BENAISSA YOUCEF")
            self.agréeeurcombobd.addItem("REZZAG SOFIANE ")
            self.agréeeurcombobd.addItem("BELBACHA M.NADIR")
            self.agréeeurcombobd.move(1133, 88)
            self.agréeeurcombobd.resize(147, 25)

            ############################################docx2pdf######################
            self.timerbd = QTimer()
            self.timerbd.timeout.connect(self.allcallculbd)
            self.timerbd.setInterval(1000)
            self.timerbd.start()

            ###########################buttons################

            self.btnsavebd = QtWidgets.QPushButton("ENREGISTRER", self.bldurtab, clicked=lambda: self.docx_bdsave())
            self.btnsavebd.move(1100, 131)
            self.btnsavebd.resize(200, 80)
            self.btnsavebd.setFont(self.font)
            self.btnsavebd.setIcon(QIcon("images/savepis.png"))
            self.btnsavebd.setIconSize(QSize(70, 80))

            self.btnprintbd = QtWidgets.QPushButton("IMPRIMER", self.bldurtab, clicked=lambda: self.printerbd())
            self.btnprintbd.move(1100, 236)
            self.btnprintbd.resize(200, 80)
            self.btnprintbd.setFont(self.font)
            self.btnprintbd.setIcon(QIcon("images/print125.png"))
            self.btnprintbd.setIconSize(QSize(70, 80))

            self.btnefacebd = QtWidgets.QPushButton("EFACER", self.bldurtab, clicked=lambda: self.clear_allbd())
            self.btnefacebd.move(1100, 341)
            self.btnefacebd.resize(200, 80)
            self.btnefacebd.setIcon(QIcon("images/eraser45877.png"))
            self.btnefacebd.setIconSize(QSize(70, 80))
            self.btnefacebd.setFont(self.font)

            # self.btncalculbd = QtWidgets.QPushButton("CALCULER", self.bldurtab,clicked=lambda :self.allcallculbd())
            # self.btncalculbd.move(1120, 525)
            # self.btncalculbd.resize(200, 80)
            # self.btncalculbd.setFont(self.font)
            # self.btncalculbd.setIcon((QIcon("images/calcul12544.png")))
            # self.btncalculbd.setIconSize(QSize(70, 80))

            self.horizontalLayout.addWidget(self.moulinwidget)
            MainWindow.setCentralWidget(self.centralwidget)
            self.statusbar = QtWidgets.QStatusBar(MainWindow)
            self.statusbar.setObjectName("statusbar")
            MainWindow.setStatusBar(self.statusbar)

            self.retranslateUi(MainWindow)
            self.moulinwidget.setCurrentIndex(1)
            QtCore.QMetaObject.connectSlotsByName(MainWindow)

        def retranslateUi(self, MainWindow):
            _translate = QtCore.QCoreApplication.translate
            MainWindow.setWindowTitle(_translate("MainWindow", "ccls relizane service qualité"))
            self.moulinwidget.setTabText(self.moulinwidget.indexOf(self.bltendretab),
                                         _translate("MainWindow", "Blé Tendre"))
            self.moulinwidget.setTabText(self.moulinwidget.indexOf(self.bldurtab), _translate("MainWindow", "Blé Dur"))


    if __name__ == "__main__":
        import sys

        app = QtWidgets.QApplication(sys.argv)
        MainWindow = QtWidgets.QMainWindow()
        ui = Moulin_Window()
        ui.mouli_window(MainWindow)
        MainWindow.show()
        sys.exit(app.exec())


except Exception as e:
    print(e)
