# %% Import Modules

from subprocess import CalledProcessError, check_output, run, STDOUT, PIPE
from tkinter import Button, Entry, filedialog, messagebox
from tkinter import StringVar, LEFT, Tk, ttk, X, CENTER
from os import getenv, listdir, makedirs, path
from datetime import datetime as dt
from shutil import copyfile
from glob import glob as gb
from re import findall
import subprocess as sp
import re

# %% Timer Start

startTime = dt.now()  # script timer

# %% Hardcoded Variables

def atoi(text):
    '''Convert text to integer if possible.'''
    return int(text) if text.isdigit() else text

def natural_keys(text):
    '''Use integers for sorting if possible.'''
    return [atoi(c) for c in re.split(r'(\d+)', text)]

JBEAM_ROOT = r"\\bosch.com\dfsrb\DfsDE\DIV\PS\PC\PS-EC\PA-SH\01_EBT\PJ\jbeam"
BINARIES_ROOT = path.join(JBEAM_ROOT, "jBEAM-stable", "Binaries")
VERSIONSTARTERS_ROOT = path.join(JBEAM_ROOT, "jBEAM-stable", "VersionStarters")
JRE_ROOT = path.join(JBEAM_ROOT, "JRE")

USERNAME = getenv('USERNAME')

JAR_LIST = listdir(BINARIES_ROOT)
JAR_ENT_LIST = [jar for jar in JAR_LIST if "jBEAM_Bosch_Enterprise-v" in jar]
JAR_REP_LIST = [jar for jar in JAR_LIST if "jBEAM_Bosch_Reporter-v" in jar]
JAR_ENT_LIST = sorted(JAR_ENT_LIST, key=natural_keys)
JAR_REP_LIST = sorted(JAR_REP_LIST, key=natural_keys)
JAR_ENT_VER = [(ent.split('-')[1]).split(".jar")[0] for ent in JAR_ENT_LIST]
JAR_REP_VER = [(rep.split('-')[1]).split(".jar")[0] for rep in JAR_REP_LIST]

VS_LIST = listdir(VERSIONSTARTERS_ROOT)
JNLP_LIST = [vs for vs in VS_LIST if vs.endswith(".jnlp")]
JNLP_ENT_LIST = [jnlp for jnlp in JNLP_LIST if "jBEAM_QuickStart_Enterprise-v" in jnlp]
JNLP_REP_LIST = [jnlp for jnlp in JNLP_LIST if "jBEAM_QuickStart_Reporter-v" in jnlp]
JNLP_ENT_VER = [(ent.split('-')[1]).split(".jnlp")[0] for ent in JNLP_ENT_LIST]
JNLP_REP_VER = [(rep.split('-')[1]).split(".jnlp")[0] for rep in JNLP_REP_LIST]
RUN_LIST = [vs for vs in VS_LIST if vs.endswith(".run")]
RUN_ENT_LIST = [run for run in RUN_LIST if "jBEAM_QuickStart_Enterprise-v" in run]
RUN_REP_LIST = [run for run in RUN_LIST if "jBEAM_QuickStart_Reporter-v" in run]
RUN_ENT_VER = [(ent.split('-')[1]).split(".jnlp")[0] for ent in RUN_ENT_LIST]
RUN_REP_VER = [(rep.split('-')[1]).split(".jnlp")[0] for rep in RUN_REP_LIST]

def atoi(text):
    '''Convert text to integer if possible.'''
    return int(text) if text.isdigit() else text

def natural_keys(text):
    '''Use integers for sorting if possible.'''
    return [atoi(c) for c in re.split(r'(\d+)', text)]

JRE_LIST = listdir(JRE_ROOT)
JRE_JDK_VER = [jdk for jdk in JRE_LIST if jdk.startswith("jdk")]
JRE_JRE_VER = [jre for jre in JRE_LIST if jre.startswith("jre")]

JRE_JDK_VER = sorted(JRE_JDK_VER, key=natural_keys)

# double start_time_double=0;
# double end_time_double=10;


GROOVY_SCRIPT = """

import com.AMS.jBEAM.*;
import com.AMS.jBEAM.FileImport_MDF;
import com.AMS.jBEAM.FileImport_MDF.MDF_FileHeader;
import com.AMS.jBEAM.AbstractDataFileImporter.LoadStatus;
import com.AMS.jBEAM.FileImport_MDF.MDF_FileHeader.MDF_DG_DataGroup;
import com.AMS.jBEAM.FileImport_MDF.MDF_FileHeader.MDF_DG_DataGroup.MDF_CG_ChannelGroup;
import com.AMS.jBEAM.FileImport_MDF.MDF_FileHeader.MDF_DG_DataGroup.MDF_CG_ChannelGroup.MDF_CN_Channel;

jBEAMParameter parameters = jB.getParameters();

String project_file = parameters.getParameter("script.ProjectFile");
String mapping_file = parameters.getParameter("script.MappingFile");
String measurement_file = parameters.getParameter("script.Measurement");
String pdf_output = parameters.getParameter("script.PdfOutput");
String excel_output = parameters.getParameter("script.ExcelOutput");
String start_time = parameters.getParameter("script.StartTime");
String end_time = parameters.getParameter("script.EndTime");
String wltp_co2u = parameters.getParameter("script.WLTPCO2U");
String wltp_co2t = parameters.getParameter("script.WLTPCO2T");
String output_mode = parameters.getParameter("script.OutputMode");

start_time_double = Double.parseDouble(start_time);
end_time_double = Double.parseDouble(end_time);

jB.openProject(project_file);

ValueInputGraph WLTPCO2U = (ValueInputGraph) jC.getComponentByName("VIG:5045");
ValueInputGraph WLTPCO2T = (ValueInputGraph) jC.getComponentByName("VIG:5045~1");
WLTPCO2U.setPublishedValue(wltp_co2u);
WLTPCO2T.setPublishedValue(wltp_co2t);

File mapping = new File(mapping_file)
File measurement = new File (measurement_file);

FileImport_MDF MeasurementImport = (FileImport_MDF) jC.getComponentByName("MDF_Measurement");
MeasurementImport.setImportFile(measurement)
MeasurementImport.setTimeRangeMode(FileImport_MDF.TimeRangeMode.MANUAL_TIME_RANGE_RESAMPLED);
MeasurementImport.setStartTimeToLoad(start_time_double);
MeasurementImport.setEndTimeToLoad(end_time_double);
MeasurementImport.setResamplingRate(10);
MeasurementImport.setCustomChannelNameUnitMappingFile(mapping);

MDF_FileHeader fileHeaderMeasurement = MeasurementImport.getFileHeader();
    for (int i = 0; i < fileHeaderMeasurement.getNumberOfAllChannelGroups(); i++) {
        MDF_DG_DataGroup dataGroupMeasurement = fileHeaderMeasurement.getDataGroup(i);
        if (dataGroupMeasurement == null) {continue;}
        for (int j = 0; j < dataGroupMeasurement.getNumberOfChannelGroups(); j++) {
            MDF_CG_ChannelGroup channelGroupMeasurement = dataGroupMeasurement.getChannelGroup(j);
            if (channelGroupMeasurement == null) {continue;}
            for (int k = 0; k < channelGroupMeasurement.getNumberOfChannelBlocks(); k++) {
                MDF_CN_Channel channelHeaderMeasurement = channelGroupMeasurement.getChannelBlock(k);
                channelHeaderMeasurement.setLoadStatus(LoadStatus.StandBy, true);
				channelHeaderMeasurement.setResultItemIndex(-1);
                channelHeaderMeasurement.setInterpolationActive(false);
                channelHeaderMeasurement.setLoadTextAsIndex(true);
                channelHeaderMeasurement.setConvDblAsStringStatus(false);}}}

MeasurementImport.invalidateHeader();
jC.validateFramework(true);

if(output_mode=="pdf") {
    def pdf_file = new File(pdf_output)
    if(pdf_file.exists()) {pdf_file.delete()}
    FileExport_PDF_Graphic pdfExport = jC.newComponent(FileExport_PDF_Graphic.class);
    pdfExport.setExportFile(new File(pdf_output));
    pdfExport.export(true);
    } else if(output_mode=="excel") {
        def excel_file = new File(excel_output)
        if(excel_file.exists()) {excel_file.delete()}
        FileExport_Excel excelExport = jC.newComponent(FileExport_Excel.class);
        excelExport.setExportFile(new File(excel_output));
        excelExport.doExport();}

jB.quitjBEAM(false, false, true);

"""

GROOVY_SCRIPT = GROOVY_SCRIPT.strip()

# %% Class Start


class Application():
    """TEST DOCSTRING."""

    def __init__(self):

        self.window = Tk()
        self.window.title('BOSCH - jBeam Launcher with StEvE')

        self.window1 = StringVar(value="opt2")
        self.window2 = StringVar(value="opt3")
        self.window3 = StringVar(value="Enterprise")
        self.window4 = StringVar(value=JAR_ENT_VER[-1])
        self.window5 = StringVar(value="Online - Network JDK")
        self.window6 = StringVar(value=JRE_JDK_VER[-1])

        self.project_path = ""
        self.mapping_path = ""
        self.measurement_path = ""

        self.make_widgets()
        self.window.mainloop()

# %%

    def make_widgets(self):

        # Frame 1

        frame1 = ttk.Frame(self.window)
        frame1.pack(fill=X)

        label1 = ttk.Label(frame1, text="jBeam Model:", width=20)
        label1.pack(side=LEFT)

        radio1 = ttk.Radiobutton(frame1, text="JNLP", variable=self.window1, value="opt1")
        radio1.pack(side=LEFT)

        radio2 = ttk.Radiobutton(frame1, text="JAR", variable=self.window1, value="opt2")
        radio2.pack(side=LEFT)

        self.window1.trace('w', self.update)

        # Frame 2

        frame2 = ttk.Frame(self.window)
        frame2.pack(fill=X)

        label2 = ttk.Label(frame2, text="Output Type:", width=20)
        label2.pack(side=LEFT)

        radio3 = ttk.Radiobutton(frame2, text="PDF", variable=self.window2, value="opt3")
        radio3.pack(side=LEFT)

        radio4 = ttk.Radiobutton(frame2, text="Excel", variable=self.window2, value="opt4")
        radio4.pack(side=LEFT)

        # Frame 3

        frame3 = ttk.Frame(self.window)
        frame3.pack(fill=X)

        label3 = ttk.Label(frame3, text="jBeam version:", width=20)
        label3.pack(side=LEFT)

        combobox1 = ttk.Combobox(frame3, textvariable=self.window3)
        combobox1["values"] = ["Enterprise", "Reporter"]
        combobox1.pack(side=LEFT)

        self.combobox2 = ttk.Combobox(frame3, textvariable=self.window4)
        self.combobox2["values"] = JNLP_ENT_VER
        self.combobox2.pack(side=LEFT)

        self.window1.trace('w', self.update_combobox2)
        self.window3.trace('w', self.update_combobox2)
        # self.window4.trace('w', self.update_combobox2)

        # Frame 4

        frame4 = ttk.Frame(self.window)
        frame4.pack(fill=X)

        label4 = ttk.Label(frame4, text="Java version:", width=20)
        label4.pack(side=LEFT)

        combobox3 = ttk.Combobox(frame4, textvariable=self.window5)
        combobox3["values"] = \
            [
                "Installed - Redhat",
                "Installed - Oracle",
                "Offline - JDK/JRE Bin",
                "Online - Network JDK",
                "Online - Network JRE"
            ]
        combobox3.pack(side=LEFT)

        self.combobox4 = ttk.Combobox(frame4, textvariable=self.window6)
        self.combobox4["values"]=JRE_JDK_VER
        self.combobox4.pack(side=LEFT)

        self.window5.trace('w', self.update_combobox4)
        # self.window6.trace('w', self.update_combobox4)

        # Frame 5

        frame5 = ttk.Frame(self.window)
        frame5.pack(fill=X)

        label5 = ttk.Label(frame5, text="Start-End Time (s) :", width=20)
        label5.pack(side=LEFT)

        self.entry1 = Entry(frame5, justify="center")
        self.entry1.pack(side=LEFT, fill=X, expand=True)

        self.entry2 = Entry(frame5, justify="center")
        self.entry2.pack(side=LEFT, fill=X, expand=True)

        self.entry1.insert(0, "0")
        self.entry2.insert(0, "2500")

        # Frame 6

        frame6 = ttk.Frame(self.window)
        frame6.pack(fill=X)

        label6 = ttk.Label(frame6, text="WLTP CO2u-CO2t :", width=20)
        label6.pack(side=LEFT)

        self.entry3 = Entry(frame6, justify="center")
        self.entry3.pack(side=LEFT, fill=X, expand=True)

        self.entry4 = Entry(frame6, justify="center")
        self.entry4.pack(side=LEFT, fill=X, expand=True)

        self.entry3.insert(0, "250")
        self.entry4.insert(0, "200")

        # Frame 7

        frame7 = ttk.Frame(self.window)
        frame7.pack(fill=X)

        label7 = ttk.Label(frame7, text="StEvE Files:", width=20)
        label7.pack(side=LEFT)

        self.button1 = Button(
            frame7,
            text="Browse",
            command=self.browse_steve,
            background="#000080",
            foreground="white"
            )
        self.button1.pack(side=LEFT, fill=X, expand=True)

        # Frame 8

        frame8 = ttk.Frame(self.window)
        frame8.pack(fill=X)

        label8 = ttk.Label(frame8, text="Measurement File:", width=20)
        label8.pack(side=LEFT)

        self.button2 = Button(
            frame8,
            text="Browse",
            command=self.browse_measurement,
            background="#000080",
            foreground="white"
            )
        self.button2.pack(side=LEFT, fill=X, expand=True)

        # Frame 9

        frame9 = ttk.Frame(self.window)
        frame9.pack(fill=X)

        button5 = Button(
            frame9,
            text="Start",
            command=self.start_process,
            background="#00563F",
            foreground="white"
            )
        button5.pack(side=LEFT, fill=X, expand=True)

    def update_combobox2(self, *args):
        """Update combobox2 when related values change."""
        if self.window1.get() == "opt1":
            if self.window3.get() == "Enterprise":
                self.combobox2['values'] = JNLP_ENT_VER
                self.combobox2.set(JNLP_ENT_VER[-1])
            elif self.window3.get() == "Reporter":
                self.combobox2['values'] = JNLP_REP_VER
                self.combobox2.set(JNLP_REP_VER[-1])
        elif self.window1.get() == "opt2":
            if self.window3.get() == "Enterprise":
                self.combobox2['values'] = JAR_ENT_VER
                self.combobox2.set(JAR_ENT_VER[-1])
            elif self.window3.get() == "Reporter":
                self.combobox2['values'] = JAR_REP_VER
                self.combobox2.set(JAR_REP_VER[-1])

    def update_combobox4(self, *args):
        """Update combobox4 when related values change."""
        if self.window5.get() == "Installed - Redhat":
            paths = gb(r"C:\Program Files\RedHat\java-*-openjdk\bin\java.exe")
            if len(paths) != 1:
                error_text = ("OpenJDK 8 Redhat cannot be found in your computer.\n"
                              "Install OpenJDK 8 Redhat with ITSP Add/Remove Software")
                messagebox.showinfo("Information", error_text)
                self.combobox4.set("Not Ready")
            else:
                try:
                    sp.check_output(["java", "-version"], stderr=sp.STDOUT)
                    self.combobox4.set("Ready")
                except:
                    error_text = ("Java is not added to system environment variables.\n"
                                  "Add folder in C:\Program Files\RedHat to S.E.V.\n"
                                  "Or javaws may not be available to use currently.")
                    messagebox.showinfo("Information", error_text)
                    self.combobox4.set("Not Ready")

        elif self.window5.get() == "Installed - Oracle":
            try:
                sp.check_output(["java", "-version"], stderr=sp.STDOUT)
                self.combobox4.set("Ready")
            except:
                error_text = ("Java is not added to system environment variables.\n"
                              "Add folder in C:\Program Files\RedHat to S.E.V.")
                messagebox.showinfo("Information", error_text)
                self.combobox4.set("Not Ready")

        elif self.window5.get() == "Offline - JDK/JRE Bin":
            if self.window1.get() == "opt1":
                info_text = "Select javaws.exe inside JDK/JRE bin folder when prompted."
                messagebox.showinfo("Information", info_text)
                self.java_path = filedialog.askopenfilename(filetypes=[("Java", "javaws.exe")])
            elif self.window1.get() == "opt2":
                info_text = "Select java.exe inside JDK/JRE bin folder when prompted."
                messagebox.showinfo("Information", info_text)
                self.java_path = filedialog.askopenfilename(filetypes=[("Java", "java.exe")])
                if self.java_path == "":
                    self.combobox4.set("Not Ready")
                else:
                    self.combobox4.set("Ready")

        elif self.window5.get() == "Online - Network JDK":
            self.combobox4['values'] = JRE_JDK_VER
            self.combobox4.set(JRE_JDK_VER[-1])

        elif self.window5.get() == "Online - Network JRE":
            self.combobox4['values'] = JRE_JRE_VER
            self.combobox4.set(JRE_JRE_VER[-1])

    def browse_steve(self):
        """sdsdsd."""
        info_text = ("You need to download StEvE template from Bosch Tools.\n"
                     "Browse .jbs file inside StEvE folder when prompted.")
        messagebox.showinfo("Information", info_text)
        self.project_path = filedialog.askopenfilename(filetypes=[("JBS File", "*.jbs")])
        if self.project_path == "":
            self.button1['text'] = "Not Ready"
            return
        info_text = ("Browse .cmap file inside StEvE folder when prompted.")
        messagebox.showinfo("Information", info_text)
        self.mapping_path = filedialog.askopenfilename(filetypes=[("JBS File", "*.cmap")])
        if self.mapping_path == "":
            self.button1['text'] = "Not Ready"
            return
        self.button1['text'] = "Ready"

    def browse_measurement(self):
        """fsdsds."""
        self.measurement_path =filedialog.askopenfilename(
            filetypes=[("Measurement", "*.mf4")])
        if self.measurement_path == "":
            self.button2['text'] = "Not Ready"
            return
        self.button2['text'] = "Ready"

    def update(self, *args):
        if self.window1.get() == "opt1":
            self.entry1.configure(state="disabled")
            self.entry2.configure(state="disabled")
            self.entry3.configure(state="disabled")
            self.entry4.configure(state="disabled")
            self.button1.configure(state="disabled")
            self.button2.configure(state="disabled")
        if self.window1.get() == "opt2":
            self.entry1.configure(state="normal")
            self.entry2.configure(state="normal")
            self.entry3.configure(state="normal")
            self.entry4.configure(state="normal")
            self.button1.configure(state="normal")
            self.button2.configure(state="normal")


    def start_process(self):
        """asdasd."""

        file_mode = self.window1.get()  # opt1 or opt2
        output_type = self.window2.get()  # opt3 or opt4
        jbeam_type = self.window3.get()  # enterprise or reporter
        jbeam_version = self.window4.get() # folder name
        java_type = self.window5.get()  # one of 5 options
        java_version = self.window6.get()  # folder name

        if file_mode == "opt2":
            start_time = self.entry1.get()
            end_time = self.entry2.get()
            wltp_co2u = self.entry3.get()
            wltp_co2t = self.entry4.get()
            steve_path = self.project_path
            mapping_path = self.mapping_path
            measurement_path = self.measurement_path

        #

        if file_mode != "opt1" and file_mode != "opt2":
            messagebox.showinfo("Information", "jBeam model has not been selected.")
            return
        if output_type != "opt3" and output_type != "opt4":
            messagebox.showinfo("Information", "Output type has not been selected.")
            return
        if jbeam_type != "Enterprise" and jbeam_type != "Reporter":
            messagebox.showinfo("Information", "jBeam version has not been selected.")
            return
        if (jbeam_version not in JNLP_ENT_VER
            and jbeam_version not in JNLP_REP_VER
            and jbeam_version not in JAR_ENT_VER
            and jbeam_version not in JAR_REP_VER):
            messagebox.showinfo("Information", "jBeam version has not been selected.")
            return
        if (java_type != "Installed - Redhat"
            and java_type != "Installed - Oracle"
            and java_type != "Offline - JDK/JRE Bin"
            and java_type != "Online - Network JDK"
            and java_type != "Online - Network JRE"):
            messagebox.showinfo("Information", "Java version has not been selected.")
            return
        if (java_version not in JRE_JDK_VER
            and java_version not in JRE_JRE_VER
            and java_version != "Ready"):
            messagebox.showinfo("Information", "Java version has not been selected.")
            return

        if file_mode == "opt2":

            if start_time == "":
                messagebox.showinfo("Information", "Start time has not been entered.")
                return
            if end_time == "":
                messagebox.showinfo("Information", "End time has not been entered.")
                return
            if wltp_co2u == "":
                messagebox.showinfo("Information", "WLTP CO2U has not been entered.")
                return
            if wltp_co2t == "":
                messagebox.showinfo("Information", "WLTP CO2T has not been entered.")
                return
            if steve_path == "":
                messagebox.showinfo("Information", "Steve file has not been selected.")
                return
            if mapping_path == "":
                messagebox.showinfo("Information", "Mapping file has not been selected.")
                return
            if measurement_path == "":
                messagebox.showinfo("Information", "Measurement file has not been selected.")
                return

        file_mode = "jnlp" if file_mode == "opt1" else "jar"
        output_type = "pdf" if output_type == "opt3" else "excel"

        #

        if java_type == "Installed - Redhat" or java_type == "Installed - Oracle":
            if file_mode == "jnlp":
                java_path = "javaws"
            if file_mode == "jar":
                java_path = "java"
        elif java_type == "Offline - JDK/JRE Bin":
            java_path = f'"{self.java_path}"'
        elif java_type == "Online - Network JDK" or java_type == "Online - Network JRE":
            if file_mode == "jnlp":
                java_path = path.join(JRE_ROOT,java_version,"bin","javaws.exe")
                java_path = f'"{java_path}"'
            if file_mode == "jar":
                java_path = path.join(JRE_ROOT,java_version,"bin","java.exe")
                java_path = f'"{java_path}"'

        if jbeam_type == "Enterprise":
            if file_mode == "jnlp":
                idx = JNLP_ENT_VER.index(jbeam_version)
                jbeam_path = path.join(VERSIONSTARTERS_ROOT, JNLP_ENT_LIST[idx])
            if file_mode == "jar":
                idx = JAR_ENT_VER.index(jbeam_version)
                jbeam_path = path.join(BINARIES_ROOT, JAR_ENT_LIST[idx])
        if jbeam_type == "Reporter":
            if file_mode == "jnlp":
                idx = JNLP_REP_VER.index(jbeam_version)
                jbeam_path = path.join(VERSIONSTARTERS_ROOT, JNLP_REP_LIST[idx])
            if file_mode == "jar":
                idx = JAR_REP_VER.index(jbeam_version)
                jbeam_path = path.join(BINARIES_ROOT, JAR_REP_LIST[idx])

        if file_mode == "jar":

            pdf_path = path.basename(measurement_path).split(".")[0] + ".pdf"
            excel_path = path.basename(measurement_path).split(".")[0] + ".xlsx"

        #

        # print(file_mode)
        # print(output_type)
        # print(jbeam_type)
        # print(jbeam_version)
        # print(java_type)
        # print(java_version)
        # print(start_time)
        # print(end_time)
        # print(wltp_co2u)
        # print(wltp_co2t)
        # print(steve_path)
        # print(mapping_path)
        # print(measurement_path)

        # print()

        # print(jbeam_path)
        # print(steve_path)
        # print(mapping_path)
        # print(measurement_path)
        # print(start_time)
        # print(end_time)
        # print(wltp_co2u)
        # print(wltp_co2t)
        # print(output_type)
        # print(pdf_path)
        # print(excel_path)

        #

        if file_mode == "jnlp":

            cmd_line = (f'{java_path}'
                        f' "{jbeam_path}"'
                        )

        if file_mode == "jar":

            cmd_line = (f'{java_path}'
                        f' -jar "{jbeam_path}"'
                        '  -ScriptURL="automation.groovy"'
                        f' -script.ProjectFile="{steve_path}"'
                        f' -script.MappingFile="{mapping_path}"'
                        f' -script.Measurement="{measurement_path}"'
                        f' -script.StartTime={start_time}'
                        f' -script.EndTime={end_time}'
                        f' -script.WLTPCO2U={wltp_co2u}'
                        f' -script.WLTPCO2T={wltp_co2t}'
                        f' -script.OutputMode={output_type}'
                        f' -script.PdfOutput="{pdf_path}"'
                        f' -script.ExcelOutput="{excel_path}"'
                        f' -Mode=Service'
                        f' -LicenseServer=10.16.1.222:4332'
                        '  -Xms4G'
                        '  -Xmx16G'
                        '  -Djava.net.preferIPv4Stack=true'
                        '  -Dsun.java2d.d3d=false'
                        '  -XX:+AggressiveOpts'
                        '  -XX:+UseParallelGC'
                        )
        with open('automation.groovy', 'w') as f:
            f.write(GROOVY_SCRIPT)

        with open('jBeamStarter.cmd', 'w') as file:
            file.write(cmd_line)

        messagebox.showinfo("Info", "jBeamStarter.cmd has been produced.\n"
                            "Run jBeamStarter.cmd at same directory.")

        self.window.destroy()


# %%

if __name__ == "__main__":

    # Make class object
    app = Application()

# %% Timer End

print("\nProgram runtime: " + str(dt.now()-startTime))
