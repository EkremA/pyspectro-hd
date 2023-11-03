# The MIT License (MIT)
#
# Copyright (c) 2022 Paul Bupe Jr, Harnett Lab
#
# Permission is hereby granted, free of charge, to any person obtaining a copy
# of this software and associated documentation files (the "Software"), to deal
# in the Software without restriction, including without limitation the rights
# to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
# copies of the Software, and to permit persons to whom the Software is
# furnished to do so, subject to the following conditions:
#
# The above copyright notice and this permission notice shall be included in
# all copies or substantial portions of the Software.
#
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
# THE SOFTWARE.

"""
A simple and fast pyqtgraph-based UI for use with Ocean Optics spectrometers.

* Author(s): Paul Bupe Jr
"""

from io import StringIO
import pandas as pd #read_csv()
import subprocess   #WindowsPowerShell

import csv
import sys
import time

import numpy as np
import pyqtgraph as pg
import pyqtgraph.exporters
from pyqtgraph.Qt import QtCore, QtGui, QtWidgets
from seabreeze.spectrometers import Spectrometer


class USB2000(QtWidgets.QMainWindow):
    def __init__(self, *args, **kwargs):
        super(USB2000, self).__init__(*args, **kwargs)
        self.integration_time_us = 200
        self.devices = ["No device connected"]
        self.spec = None
        self.initialized = False
        self.wavelengths = []
        self.intensities = []
        self.averaged_intensities = []
        self.scans_to_average = 4
        self.capturing = False

        self.build_main_ui()
        self.build_toolbar()
        self.build_plot()

        self.start_time = time.time()
        self.timer = QtCore.QTimer()
        self.timer.timeout.connect(self.update)
        self.timer.start(30)  # USB scans into memory every 13ms
        self.show()

    def update(self):
        if self.capturing:
            if self.initialized:
                self.read_spectrum()
                self.plot.setData(self.wavelengths, self.averaged_intensities)

    def init_spectrometer(self):
        try:
            """#TODO: outsource this to a function and have detach as well! detach first always by default when starting program (for now, but maybe later in destructor method in case somebody x's the window and it's still attached, we dont want USB eject related loss or damage...)
            powerShellCommand = ["/mnt/c/Windows/System32/WindowsPowerShell/v1.0/powershell.exe", 
                                 "usbipd wsl list", 
                                 " | ", 
                                 "Select-String -Pattern \"Ocean Optics\""]
            print("Finding all connected USB devices with \"Ocean Optics\" in name...")
            print(powerShellCommand)
            powerShellMessage = subprocess.run(powerShellCommand, capture_output=True)
            listUSBIPDDevices = powerShellMessage
            listUSBIPDDevices = listUSBIPDDevices.stdout.decode().strip().split("  ")   #hacky string manipulation...
            listUSBIPDDevices = list(filter(None, listUSBIPDDevices))                   #hacky string manipulation...
            firstHitBUSID = listUSBIPDDevices[0]
            print(firstHitBUSID)
            time.sleep(3)
            
            powerShellCommand = ["/mnt/c/Windows/System32/WindowsPowerShell/v1.0/powershell.exe", 
                                 "usbipd wsl  detach --busid {} --distribution Ubuntu".format(str(firstHitBUSID))]
            print("Disconnecting first Spectrometer of list...")
            print(powerShellCommand)
            powerShellMessage = subprocess.run(powerShellCommand, capture_output=True)
            time.sleep(3) # buffer time to let powershell command get done (subprocess.run oddly not blocking as command... we hack around)

            powerShellCommand = ["/mnt/c/Windows/System32/WindowsPowerShell/v1.0/powershell.exe", 
                                 "usbipd wsl  attach --busid {} --distribution Ubuntu".format(str(firstHitBUSID))]
            print("Reconnecting first Spectrometer of list...")
            print(powerShellCommand)
            powerShellMessage = subprocess.run(powerShellCommand, capture_output=True)
            #TODO: You have to press Connect two times such that a device can be found. Is there a lag here? run() should be blocking
            #TODO: sudo apt install linux-tools-virtual hwdata && sudo update-alternatives --install /usr/local/bin/usbip usbip `ls /usr/lib/linux-tools/*/usbip | tail -n1` 20 #Before every start I would suggest because sometimes WSL breaks possibly because of update
            # print("Made it to sleep")
            print("Connecting to Ocean Optics Spectrometer...")
            time.sleep(3) # buffer time to let powershell command get done (subprocess.run oddly not blocking as command... we hack around)
            #print("Woke up")"""
            self.spec = Spectrometer.from_first_available()
            print("Device initialized: " + self.spec.model)
            self.spec.integration_time_micros(self.integration_time_us)
            self.wavelengths = self.spec.wavelengths()
            self.plot_widget.setXRange(
                self.wavelengths[0], self.wavelengths[-1], padding=0
            )

            self.initialized = True
            self.status_label.setText("Device initialized: " + self.spec.model)

            # disable connection button
            self.connect_button.setEnabled(False)
            self.start_capture_button.setEnabled(True)
        except Exception as e:
            print(e)
            self.status_label.setText("No device detected")
            return

    def read_spectrum(self):

        self.intensities = []

        for i in range(self.scans_to_average):
            try:
                intensities = self.spec.intensities()
                self.intensities.append(intensities)
            except Exception as e:
                print(e)

        # If list not empty, average
        if len(self.intensities) > 0:
            self.averaged_intensities = np.mean(self.intensities, axis=0)
            self.intensities = []
        # self.averaged_intensities = np.mean(self.intensities, axis=0)

    def update_device_list(self):
        self.devices_dropdown.currentIndexChanged.disconnect()
        self.devices_dropdown.clear()
        self.devices_dropdown.addItems(self.devices)
        self.devices_dropdown.currentIndexChanged.connect(self.update_device_list)

    def update_integration_time(self):
        try:
            self.integration_time_us = int(self.integration_time_edit.text())
            self.spec.integration_time_micros(self.integration_time_us)
        except Exception as e:
            print(e)

    def update_scans_to_average(self):
        try:
            self.scans_to_average = int(self.scans_to_average_edit.text())
        except Exception as e:
            print(e)

    def start_capture(self):
        self.capturing = True
        # disable start button
        self.start_capture_button.setEnabled(False)
        self.stop_capture_button.setEnabled(True)

    def stop_capture(self):
        self.capturing = False

        # enable start button
        self.start_capture_button.setEnabled(True)
        self.stop_capture_button.setEnabled(False)
        # clear plot
        # self.plot.clear()

    def export_screenshot(self):
        exporter = pg.exporters.ImageExporter( self.plot_widget.scene() )
        fileName = self.spec.model+"_"+self.spec.serial_number+"_"+time.strftime("%d-%b-%Y_%H-%M-%S")
        #pg.Qt.QGuiApplication.processEvents()
        exporter.export("{}.png".format(fileName))

    def export_csv(self):
        if self.intensities is None:
            self.status_label.setText("ERROR: No data to export")
            return

        # get file name
        filename = QtWidgets.QFileDialog.getSaveFileName(
            self, "Export Data", "", "CSV Files (*.csv)"
        )

        filename = filename[0]
        if not filename:
            # update status label
            self.status_label.setText("ERROR: No file selected")
            return

        # write data to file
        with open(filename, "w", newline="") as f:
            writer = csv.writer(f)
            writer.writerow(["Wavelength", "Intensity"])
            for i in range(len(self.wavelengths)):
                writer.writerow([self.wavelengths[i], self.averaged_intensities[i]])

        self.status_label.setText("INFO: Data exported to %s" % filename)

    def import_csv(self):
        #Take data in with pandas
        #Plot it and make sure update plot does not get rid of this imported data
        dlg = pg.FileDialog()
        if dlg.exec_():
            filenames = dlg.selectedFiles()
            
            self.imported_wavelengths, self.imported_averaged_intensities = np.genfromtxt(filenames[0],
                                                                                          delimiter=",",
                                                                                          skip_header=1,
                                                                                          unpack=True)
            self.imported_data.setData(np.zeros(1)) #one imported plot at a time
            self.imported_data.setData(self.imported_wavelengths, self.imported_averaged_intensities)
            print(self.plot_widget.getPlotItem())

    def mouse_moved(self, event):
        pos = event[0]
        if self.plot_widget.sceneBoundingRect().contains(pos):
            mouse_point = self.plot_widget.plotItem.vb.mapSceneToView(pos)
            """index = int(mouse_point.x())
            if index > 0 and index < len(self.wavelengths):
                self.status_label.setText(
                    "Wavelength: {:.2f} nm".format(self.wavelengths[index])
                )"""
            self.v_line.setPos(mouse_point.x())
            self.h_line.setPos(mouse_point.y())
            # self.label_cursorpos.setText("x: {:.2f}\ny: {:.2f}".format(mouse_point.x(),mouse_point.y()))
            # self.label_cursorpos.setPos(QtCore.QPointF(min(self.plot_widget.getViewBox().viewRange()[0]),
            #                                         	max(self.plot_widget.getViewBox().viewRange()[1])))

    def build_main_ui(self):
        self.main_layout = QtWidgets.QGridLayout()
        self.setCentralWidget(QtWidgets.QWidget(self))
        self.centralWidget().setLayout(self.main_layout)
        self.setWindowTitle("Ocean Insight Spectrometer")
        self.setStyleSheet(
            """
        QMainWindow {background-color: black;}
        QPushButton {
            background-color: rgb(197, 198, 199); 
            border: 0px;
            border-radius: 2px;
            padding: 5px;}
        QPushButton:pressed {
            background-color: rgb(150, 150, 150);}
        QToolBar {
            background-color: rgb(25, 25, 25);
            color: rgb(190, 190, 190);
            border: 0px;
            padding: 5px;
            spacing: 5px;}
        QToolBar::separator {
            color: rgb(190, 190, 190);
            }
              """
        )
        self.resize(1300, 600)
        pg.setConfigOptions(antialias=True)

    def build_plot(self):
        self.plot_widget = pg.PlotWidget()
        # self.plot_widget.setBackground("w")
        self.plot_widget.setLabel("left", "Intensity", units="counts", unitPrefix=None)
        self.plot_widget.getAxis("left").enableAutoSIPrefix(False)

        self.plot_widget.setLabel("bottom", "Wavelength", units="nm", unitPrefix=None)
        self.plot_widget.getAxis("bottom").enableAutoSIPrefix(False)

        self.plot_widget.setXRange(330, 1025)
        self.plot_widget.setYRange(0, 5000)
        self.plot_widget.showGrid(x=True, y=True, alpha=0.5)
        self.plot_widget.setMouseEnabled(x=False, y=False)

        
        # cursor v and h line
        self.v_line = pg.InfiniteLine(angle=90, movable=False)
        self.plot_widget.addItem(self.v_line, ignore_bounds=True)
        self.h_line = pg.InfiniteLine(angle=0, movable=False)
        self.plot_widget.addItem(self.h_line, ignore_bounds=True)
        
        # placeholder for imported csv
        self.imported_data = pg.PlotDataItem(np.zeros(1))#genfromtxt('my_file.csv') #TODO: pandas import csv here
        self.plot_widget.addItem(self.imported_data)
        
        
        self.label_cursorpos = pg.TextItem('', **{'color': '#FFF'})
        self.plot_widget.addItem(self.label_cursorpos)
        self.label_cursorpos.setPos(QtCore.QPointF(min(self.plot_widget.getViewBox().viewRange()[0]),
                                                   min(self.plot_widget.getViewBox().viewRange()[1])))
        print(self.plot_widget.getViewBox().viewRange())
        self.plot = self.plot_widget.plot(np.zeros(1), pen=(0, 255, 0))
        self.proxy = pg.SignalProxy(
            self.plot_widget.scene().sigMouseMoved, rateLimit=60, slot=self.mouse_moved
        )
        self.main_layout.addWidget(self.plot_widget, 0, 0, 1, 1)

    def build_toolbar(self):
        self.toolbar = self.addToolBar("Toolbar")
        self.toolbar.setMovable(False)
        self.toolbar.setFloatable(False)

        # int validation
        int_validator = QtGui.QIntValidator()

        self.connect_button = QtWidgets.QPushButton("Connect")
        self.connect_button.clicked.connect(self.init_spectrometer)
        self.toolbar.addWidget(self.connect_button)

        integration_time_label = QtWidgets.QLabel("Integration time (us):")
        integration_time_label.setStyleSheet(
            "color: rgb(190, 190, 190); margin-left: 10px;"
        )
        self.toolbar.addWidget(integration_time_label)

        self.integration_time_edit = QtWidgets.QLineEdit(str(self.integration_time_us))
        self.integration_time_edit.setFixedWidth(60)
        self.integration_time_edit.setValidator(int_validator)
        self.integration_time_edit.setAlignment(QtCore.Qt.AlignRight)
        self.integration_time_edit.editingFinished.connect(self.update_integration_time)
        self.toolbar.addWidget(self.integration_time_edit)

        scans_to_average_label = QtWidgets.QLabel("Scans to average:")
        scans_to_average_label.setStyleSheet(
            "color: rgb(190, 190, 190); margin-left: 10px;"
        )
        self.toolbar.addWidget(scans_to_average_label)
        
        self.scans_to_average_edit = QtWidgets.QLineEdit(str(self.scans_to_average))
        self.scans_to_average_edit.setFixedWidth(40)
        self.scans_to_average_edit.setValidator(int_validator)
        self.scans_to_average_edit.setAlignment(QtCore.Qt.AlignRight)
        self.scans_to_average_edit.editingFinished.connect(self.update_scans_to_average)
        self.toolbar.addWidget(self.scans_to_average_edit)

        # Add status label
        self.status_label = QtWidgets.QLabel("STATUS: No device connected")
        self.status_label.setStyleSheet("color: white; margin-left: 10px;")
        self.toolbar.addWidget(self.status_label)

        # create spacer widget
        spacer = QtWidgets.QWidget()
        spacer.setSizePolicy(
            QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding
        )
        self.toolbar.addWidget(spacer)

        self.start_capture_button = QtWidgets.QPushButton("Start capture")
        self.start_capture_button.setFixedWidth(75)
        self.start_capture_button.clicked.connect(self.start_capture)
        self.start_capture_button.setEnabled(False)
        self.toolbar.addWidget(self.start_capture_button)

        self.stop_capture_button = QtWidgets.QPushButton("Stop capture")
        self.stop_capture_button.setFixedWidth(75)
        self.stop_capture_button.clicked.connect(self.stop_capture)
        self.stop_capture_button.setEnabled(False)
        self.toolbar.addWidget(self.stop_capture_button)
        
        self.toolbar.addSeparator()

        self.screenshot_button = QtWidgets.QPushButton("Screenshot PNG")
        self.screenshot_button.setFixedWidth(75)
        self.screenshot_button.clicked.connect(self.export_screenshot)
        self.toolbar.addWidget(self.screenshot_button)

        self.import_button = QtWidgets.QPushButton("Import CSV")
        self.import_button.setFixedWidth(75)
        self.import_button.clicked.connect(self.import_csv)
        self.toolbar.addWidget(self.import_button)


        self.export_button = QtWidgets.QPushButton("Export CSV")
        self.export_button.setFixedWidth(75)
        self.export_button.clicked.connect(self.export_csv)
        self.toolbar.addWidget(self.export_button)


def test():
    spec = Spectrometer.from_first_available()
    while True:
        print(spec.wavelengths())
        print(spec.intensities())
        # devices = list_devices()
        time.sleep(0.1)


def main():
    global app 
    app = QtWidgets.QApplication(sys.argv)
    usb2000 = USB2000()
    sys.exit(app.exec_())


if __name__ == "__main__":
    # test()
    main()
