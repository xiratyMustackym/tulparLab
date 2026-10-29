import math
import os
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
from itertools import product

# Betiğin disk üzerindeki gerçek klasör yolunu sabitliyoruz
BASE_DIR = os.path.dirname(os.path.abspath(__file__))


class TulparHub:
    def __init__(self, root):
        self.mainPage = root
        self.mainPage.title("tulparLab v1.0")
        self.mainPage.geometry("1200x700")
        self.mainPage.configure(bg="#2b2b2b")

        # TTK Stil Yapılandırması
        self.style = ttk.Style()
        self.style.theme_use('clam')
        
        # Treeview Stil Ayarları
        self.style.configure("Treeview", 
                            background="white", 
                            foreground="black", 
                            fieldbackground="white",
                            font=('Arial', 10))
        self.style.configure("Treeview.Heading", 
                            background="#e1e1e6", 
                            foreground="black", 
                            font=('Arial', 10, 'bold'))
        self.style.map("Treeview", background=[('selected', '#0078d7')], foreground=[('selected', 'white')])

        # Ana Taşıyıcı Çerçeve
        self.container = tk.Frame(self.mainPage, bg="#2b2b2b")
        self.container.pack(fill=tk.BOTH, expand=True)

        # Gerekli Temel Klasörleri Oluştur
        self.makeDirs()

        # Modülleri Başlat
        self.tensileAnalysisApp = TensileAnalyzer(self)
        self.electricalAnalysisApp = ElectricalAnalyzer(self)
        self.xrdAnalysisApp = XRDAnalyzer(self)

        # Hub Ekranını Göster
        self.setupHubPage()

    def makeDirs(self):
        os.makedirs(os.path.join(BASE_DIR, "Tensile Tests"), exist_ok=True)
        os.makedirs(os.path.join(BASE_DIR, "Electrical Tests"), exist_ok=True)
        os.makedirs(os.path.join(BASE_DIR, "XRD Tests"), exist_ok=True)

    def setupHubPage(self):
        self.hubPage = tk.Frame(self.container, bg="#2b2b2b")
        # 1. Başlık Çerçevesi (Arka planı beyaz, genişlik/yükseklik vermiyoruz, otomatik sığacak)
        titleFrame = tk.Frame(self.hubPage, bg="#2b2b2b")
        # fill="x" ekleyerek çerçevenin yatayda güzelce yayılmasını sağlayabilirsiniz (isteğe bağlı)
        titleFrame.pack(pady=40, padx=40,expand=False)

        # 2. Ana Başlık Yazısı (titleFrame içine yerleşiyor)
        # bg="white" ekledik ki çerçevenin beyaz rengiyle bütünleşsin. 
        # pady değerini 50'den 10-20 civarına çekmek daha dengeli bir görünüm sağlar.
        tk.Label(
            titleFrame, 
            text="Tulpar Lab", 
            font=("Arial", 26, "bold"), 
            fg="#E0E0E0", 
            bg="#333333",
            highlightbackground="#FFFFFF",  # Kenarlığın RENGİ
            highlightthickness=3,
            width = 15
        ).pack(pady=10)

        # 3. Alt Başlık Yazısı (Ana sayfaya yerleşiyor)
        tk.Label(
            self.hubPage, 
            text="Select a module to launch:", 
            font=("Arial", 14), 
            bg="#2b2b2b", 
            fg="white"
        ).pack(pady=10)


        # Tensile Launch Button
        btnTensileLaunch = tk.Button(
            self.hubPage, 
            text="Tensile Test Analyzer ➔", 
            command=self.tensileAnalysisApp.showTensilePage, 
            bg="light gray", fg="black", 
            font=("Arial", 14, "bold"), 
            width=40
        )
        btnTensileLaunch.pack(pady=15)

        # Electrical Launch Button
        btnElectricalLaunch = tk.Button(
            self.hubPage,
            text="4-Point Probe & I-V Characterization Analyzer ➔",
            command=self.electricalAnalysisApp.showElectricalPage, 
            bg="light gray", fg="black",
            font=("Arial", 14, "bold"),
            width=40
        )
        btnElectricalLaunch.pack(pady=15)

        # XRD Launch Button
        btnXRDLaunch = tk.Button(
            self.hubPage,
            text="X-Ray Diffraction (XRD) Crystallography Analyzer ➔",
            command=self.xrdAnalysisApp.showXRDPage, 
            bg="light gray", fg="black",
            font=("Arial", 14, "bold"),
            width=40
        )
        btnXRDLaunch.pack(pady=15)

        self.hubPage.pack(fill=tk.BOTH, expand=True)


class TensileAnalyzer:
    def __init__(self, referenceOfHub):
        self.hub = referenceOfHub

        self.currentFilePath = None       
        self.dataFrameOfMaterials = None
        self.materialNames = ["No Database Loaded"]
        
        self.localVariableOfMaterial = tk.StringVar(self.hub.container)
        self.localVariableOfMaterial.set(self.materialNames[0])

        self.setupTensilePage()
        self.setupGraphPage()  
        self.checkAndCreateDatabase()

    def checkAndCreateDatabase(self):
        tensileFolder = "Tensile Tests"
        os.makedirs(os.path.join(BASE_DIR, tensileFolder), exist_ok=True)
        
        self.currentFilePath = os.path.join(BASE_DIR, tensileFolder, "tensilematerials.xlsx")
        
        if not os.path.exists(self.currentFilePath):
            try:
                defaultData = {
                    "MaterialName": ["Structural Steel", "Aluminum Alloy (6061-T6)", "Titanium Grade 5"],
                    "ModuleofYoung": [210000, 70000, 113800],
                    "YieldStrength": [250, 270, 880],
                    "UltimateTensileStrength": [400, 310, 950],
                    "StrengthCoefficient": [680, 410, 1250],
                    "StrainHardeningExponent": [0.20, 0.08, 0.10],
                    "FractureStress": [300, 200, 700],
                    "FractureStrain": [0.25, 0.17, 0.12]
                }
                defaultDF = pd.DataFrame(defaultData)
                defaultDF.set_index("MaterialName", inplace=True)
                defaultDF.to_excel(self.currentFilePath)
            except Exception as e:
                print(f"Could not create default tensile database: {e}")

        try:
            self.dataFrameOfMaterials = pd.read_excel(self.currentFilePath)
            self.dataFrameOfMaterials.columns = [
                col.strip().lower().replace(" ", "").replace("-", "").replace("_", "").replace("/", "") 
                for col in self.dataFrameOfMaterials.columns
            ]
            if "materialname" in self.dataFrameOfMaterials.columns:
                self.dataFrameOfMaterials.set_index("materialname", inplace=True)

            self.materialNames = self.dataFrameOfMaterials.index.tolist()
            menu = self.menuOfMaterial["menu"]
            menu.delete(0, "end")
            for name in self.materialNames:
                menu.add_command(label=name, command=tk._setit(self.localVariableOfMaterial, name))
            self.localVariableOfMaterial.set(self.materialNames[0])
        except Exception as e:
            print(f"Could not load tensile database: {e}")

    def showHubPage(self):
        self.tensilePage.pack_forget()
        self.graphPage.pack_forget()
        self.hub.hubPage.pack(fill=tk.BOTH, expand=True)

    def showTensilePage(self):
        self.hub.hubPage.pack_forget()
        self.graphPage.pack_forget()
        self.tensilePage.pack(fill=tk.BOTH, expand=True)

    def setupTensilePage(self):
        self.tensilePage = tk.Frame(self.hub.container, bg="#2b2b2b")

        topNav = tk.Frame(self.tensilePage, bg="#808080", height=45)
        topNav.pack(fill=tk.X, side=tk.TOP)

        btnBackToHub = tk.Button(topNav, text="⬅ Back to Hub", font=("Arial", 10, "bold"), bg="white", fg="black", command=self.showHubPage)
        btnBackToHub.pack(side=tk.LEFT, padx=15, pady=6)

        tk.Label(topNav, text="Tensile Test Analyzer Module", font=("Arial", 12, "bold"), bg="#808080", fg="white").pack(side=tk.LEFT, padx=15)

        contentFrame = tk.Frame(self.tensilePage, bg="#2b2b2b")
        contentFrame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        self.leftControlFrame = tk.Frame(contentFrame, bg="#2b2b2b")
        self.leftControlFrame.pack(side=tk.LEFT, fill=tk.Y, padx=15, pady=10)

        self.mainLabel = tk.Label(self.leftControlFrame, text="Choose a Material:", font=("Arial", 14, "bold"), bg="#2b2b2b", fg="white")
        self.mainLabel.grid(row=0, column=0, columnspan=2, pady=(5, 5))
    
        self.localVariableOfMaterial = tk.StringVar(self.leftControlFrame)
        self.localVariableOfMaterial.set(list(self.materialNames)[0])
        self.menuOfMaterial = tk.OptionMenu(self.leftControlFrame, self.localVariableOfMaterial, *self.materialNames)
        self.menuOfMaterial.config(width=18, bg="white", fg="black", activebackground="white", activeforeground="black")
        self.menuOfMaterial.grid(row=1, column=0, columnspan=2, pady=(0, 10))
    
        self.parameters = [
            ("youngEntry", "E", "Module of Young (E):", 2), 
            ("yieldStrengthEntry", "yieldStrength", "Yield of Strength:", 3), 
            ("UTSEntry", "UTS", "UTS:", 4),
            ("strengthCoefficientEntry", "K", "Strength Coefficient (K):", 5), 
            ("strainHardeningExponentEntry", "n", "Strain Hardening (n):", 6),
            ("fractureStressEntry", "fractureStress", "Fracture Stress:", 7), 
            ("fractureStrainEntry", "fractureStrain", "Fracture Strain:", 8),
            ("diameterEntry", "diameter", "Diameter (mm):", 9), 
            ("lengthEntry", "length", "Length (mm):", 10)
        ]
        
        for attrName, attrVar, labelText, rowIdx in self.parameters:
            tk.Label(self.leftControlFrame, text=labelText, font=("Arial", 9, "bold"), bg="#2b2b2b", fg="white").grid(row=rowIdx, column=0, sticky="w", padx=5, pady=3)
            entry = tk.Entry(self.leftControlFrame, width=12, bg="white", fg="black", insertbackground="black")
            entry.grid(row=rowIdx, column=1, padx=5, pady=3)
            setattr(self, attrName, entry)

        self.diameterEntry.insert(0, "12.5")
        self.lengthEntry.insert(0, "50.0")

        self.isLowCarbonSteel = tk.BooleanVar(value=False)
        self.steelCheckButton = tk.Checkbutton(
            self.leftControlFrame, text="Low Carbon Steel (Yield Drop)", 
            variable=self.isLowCarbonSteel, font=("Arial", 9, "bold"), 
            fg="#ff6b6b", bg="#2b2b2b", activebackground="#2b2b2b", selectcolor="white"
        )
        self.steelCheckButton.grid(row=11, column=0, columnspan=2, pady=5)
    
        self.uploadButton = tk.Button(self.leftControlFrame, text="Load Excel", command=self.loadExternalExcel, bg="blue", fg="black", font=("Arial", 9, "bold"), width=13)
        self.uploadButton.grid(row=12, column=0, padx=2, pady=3)

        self.deleteButton = tk.Button(self.leftControlFrame, text="Delete Selected", command=self.deleteSelectedMaterial, bg="red", fg="black", font=("Arial", 9, "bold"), width=20)
        self.deleteButton.grid(row=12, column=1, padx=2, pady=5)

        self.runButton = tk.Button(self.leftControlFrame, text="-Run Simulation-", bg="green", fg="black", font=("Arial", 14, "bold"), command=self.runSimulator, width=40)
        self.runButton.grid(row=13, column=0, columnspan=2, pady=8)
    
        tk.Label(self.leftControlFrame, text="New Material Name:", font=("Arial", 9, "bold"), fg="#87ceeb", bg="#2b2b2b").grid(row=14, column=0, sticky="w", padx=5, pady=(10, 2))
        self.newMaterialNameEntry = tk.Entry(self.leftControlFrame, width=15, bg="white", fg="black", insertbackground="black")
        self.newMaterialNameEntry.grid(row=14, column=1, padx=5, pady=(10, 2))
    
        self.saveButton = tk.Button(self.leftControlFrame, text="Save the New Material", bg="orange", fg="black", font=("Arial", 9, "bold"), command=self.saveNewMaterialToExcel, width=20)
        self.saveButton.grid(row=15, column=0, columnspan=2, padx=2, pady=5)

        self.compareButton = tk.Button(self.leftControlFrame, text="Compare CSVs", command=self.compareAndShowResults, bg="purple", fg="black", font=("Arial", 9, "bold"), width=30)
        self.compareButton.grid(row=16, column=0, columnspan=2, pady=3)

        self.localVariableOfMaterial.trace_add("write", self.updateEntries)

        self.rightResultFrame = tk.Frame(contentFrame, relief="sunken", bd=1, bg="#2b2b2b")
        self.rightResultFrame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=15, pady=10)

        tk.Label(self.rightResultFrame, text="Analysis Results", font=("Arial", 14, "bold"), fg="#fbfbfb", bg="#2b2b2b").pack(pady=10)
        
        self.resultTree = ttk.Treeview(self.rightResultFrame, columns=("Value", "Unit"), selectmode="none", height=23)
        self.resultTree.heading("#0", text="Property")
        self.resultTree.heading("Value", text="Value")
        self.resultTree.heading("Unit", text="Unit")
        self.resultTree.column("#0", width=250, anchor="w")
        self.resultTree.column("Value", width=110, anchor="center")
        self.resultTree.column("Unit", width=80, anchor="center")
        
        self.resultTree.tag_configure("better", foreground="green")
        self.resultTree.tag_configure("worse", foreground="red")
        self.resultTree.tag_configure("equal", foreground="orange")
        
        self.resultTree.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        self.btnGraphForward = tk.Button(self.rightResultFrame, text="Show Graph 📈 ➔", font=("Arial", 11, "bold"), bg="white", fg="black", command=self.showGraphPage)
        self.btnGraphForward.pack(anchor="e", padx=10, pady=10)

    def setupGraphPage(self):
        self.graphPage = tk.Frame(self.hub.container, bg="#2b2b2b")
        
        graphTopNav = tk.Frame(self.graphPage, bg="white", height=45)
        graphTopNav.pack(fill=tk.X, side=tk.TOP)

        btnBackToResults = tk.Button(graphTopNav, text="Back to Results", font=("Arial", 10, "bold"), bg="gray", fg="black", command=self.showTensilePage)
        btnBackToResults.pack(side=tk.LEFT, padx=15, pady=6)

        saveGraphBtn = tk.Button(graphTopNav, text="Save Graph Image", font=("Arial", 10, "bold"), bg="orange", fg="black", command=self.saveGraphImage)
        saveGraphBtn.pack(side=tk.RIGHT, padx=15, pady=6)

        self.graphCanvasFrame = tk.Frame(self.graphPage, bg="white")
        self.graphCanvasFrame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

    def showGraphPage(self):
        if hasattr(self, 'comparisonData') and self.comparisonData:
            self.showComparisonGraph()
            return

        if not hasattr(self, 'fig'):
            messagebox.showwarning("Warning", "Please run a simulation or compare CSV files first!")
            return
            
        self.tensilePage.pack_forget()
        self.graphPage.pack(fill=tk.BOTH, expand=True)

        for widget in self.graphCanvasFrame.winfo_children():
            widget.destroy()

        canvas = FigureCanvasTkAgg(self.fig, master=self.graphCanvasFrame)
        canvas.draw()
        
        toolbar = NavigationToolbar2Tk(canvas, self.graphCanvasFrame)
        toolbar.update()
        
        canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

    def startLiveAnimation(self):
        self.tensilePage.pack_forget()
        self.graphPage.pack(fill=tk.BOTH, expand=True)
        
        if hasattr(self, 'fig'):
            self.fig.clf() # Eski figürün içeriğini temizle
       
       
        for widget in self.graphCanvasFrame.winfo_children():
            widget.destroy()

        self.fig = Figure(figsize=(10, 5), dpi=100)
        self.ax1 = self.fig.add_subplot(121) 
        self.ax2 = self.fig.add_subplot(122) 

        self.ax1.set_title("Graph of Force-Extension (Live)")
        self.ax1.set_xlabel("Extension (mm)")
        self.ax1.set_ylabel("Force (N)")
        self.ax1.grid(True)

        self.ax2.set_title("Graph of Stress-Strain (Live)")
        self.ax2.set_xlabel("Strain (mm/mm)")
        self.ax2.set_ylabel("Stress (MPa)")
        self.ax2.grid(True)

        self.canvas = FigureCanvasTkAgg(self.fig, master=self.graphCanvasFrame)
        self.canvas.draw()
        
        self.toolbar = NavigationToolbar2Tk(self.canvas, self.graphCanvasFrame)
        self.toolbar.update()
        
        self.canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

        self.updateAnimation()

    def updateAnimation(self):
        totalSteps = len(self.extensionList)
        
        if self.animIndex < totalSteps:
            stepChunk = max(1, totalSteps // 150) 
            currentEnd = min(self.animIndex + stepChunk, totalSteps)

            extSlice = self.extensionList[:currentEnd]
            forceSlice = self.forceList[:currentEnd]
            strainSlice = self.strainList[:currentEnd]
            stressSlice = self.stressList[:currentEnd]

            self.ax1.clear()
            self.ax2.clear()

            self.ax1.plot(extSlice, forceSlice, color="red", linewidth=2)
            self.ax1.set_title("Graph of Force-Extension (Live)")
            self.ax1.set_xlabel("Extension (mm)")
            self.ax1.set_ylabel("Force (N)")
            self.ax1.grid(True)

            self.ax2.plot(strainSlice, stressSlice, color="blue", linewidth=2)
            self.ax2.set_title("Graph of Stress-Strain (Live)")
            self.ax2.set_xlabel("Strain (mm/mm)")
            self.ax2.set_ylabel("Stress (MPa)")
            self.ax2.grid(True)

            self.canvas.draw()

            self.animIndex = currentEnd
            self.hub.mainPage.after(5, self.updateAnimation)
        else:
            self.ax1.plot(self.extensionList[self.placeOfYieldPoint], self.forceList[self.placeOfYieldPoint], 'go', markersize=6)
            self.ax1.annotate("Yield Point", xy=(self.extensionList[self.placeOfYieldPoint], self.forceList[self.placeOfYieldPoint]), xytext=(5,5), textcoords="offset points", fontsize=9, color="green", fontweight="bold")        
            
            self.ax2.plot(self.strainList[self.placeOfYieldPoint], self.stressList[self.placeOfYieldPoint], 'go', markersize=6)
            self.ax2.annotate("Yield Point", xy=(self.strainList[self.placeOfYieldPoint], self.stressList[self.placeOfYieldPoint]), xytext=(5,5), textcoords="offset points", fontsize=9, color="green", fontweight="bold")

            self.ax1.plot(self.extensionList[self.placeOfMaxStress], self.forceList[self.placeOfMaxStress], 'go', markersize=6)
            self.ax1.annotate("UTS Point", xy=(self.extensionList[self.placeOfMaxStress], self.forceList[self.placeOfMaxStress]), xytext=(5,5), textcoords="offset points", fontsize=9, color="green", fontweight="bold")
            
            self.ax2.plot(self.strainList[self.placeOfMaxStress], self.stressList[self.placeOfMaxStress], 'go', markersize=6)
            self.ax2.annotate("UTS Point", xy=(self.strainList[self.placeOfMaxStress], self.stressList[self.placeOfMaxStress]), xytext=(5,5), textcoords="offset points", fontsize=9, color="green", fontweight="bold")

            self.ax1.plot(self.extensionList[-1], self.forceList[-1], 'gx', markersize=7)
            self.ax1.annotate("Fracture Point", xy=(self.extensionList[-1], self.forceList[-1]), xytext=(5,5), textcoords="offset points", fontsize=9, color="green", fontweight="bold")
            
            self.ax2.plot(self.strainList[-1], self.stressList[-1], 'gx', markersize=7)
            self.ax2.annotate("Fracture Point", xy=(self.strainList[-1], self.stressList[-1]), xytext=(5,5), textcoords="offset points", fontsize=9, color="green", fontweight="bold")

            self.textOfToughness = f"Toughness = {self.toughnessOfMaterial:.4f} MJ/m³"
            self.ax2.text(0.70, -0.10, self.textOfToughness, transform=self.ax2.transAxes, fontsize=10, bbox=dict(facecolor="white", edgecolor="black", alpha=0.7))

            self.canvas.draw()
            messagebox.showinfo("Success", f"Simulation completed and results saved to:\n{self.resultFolder}")

    def testEngine(self, defDiameter, defLength, E, yieldStrength, UTS, K, n, fractureStress, fractureStrain, isLowCarbon):
        self.defCrossSectionalArea = (defDiameter ** 2) * math.pi / 4
        
        if isLowCarbon and yieldStrength > 0:
            upperYield = yieldStrength * 1.07
            stressElastic = np.linspace(0, upperYield, 600)
            strainElastic = stressElastic / E

            lowerYield = yieldStrength * 0.96
            strainDrop = np.linspace(strainElastic[-1], strainElastic[-1] + 0.002, 100)
            stressDrop = np.linspace(upperYield, lowerYield, 100)

            strainPlateau = np.linspace(strainDrop[-1], strainDrop[-1] + 0.015, 400)
            baseWave = np.sin(np.linspace(0, 25 * np.pi, 400)) * (yieldStrength * 0.0015)
            noise = np.random.uniform(-1, 1, 400) * (yieldStrength * 0.0015)
            oscillations = baseWave + noise

            stressPlateau = lowerYield + oscillations

            stressPlastic = np.linspace(stressPlateau[-1], UTS, 1500)
            strainPlastic = strainPlateau[-1] + ((stressPlastic - stressPlateau[-1]) / E) + ((stressPlastic / K)**(1/n) - (stressPlateau[-1] / K)**(1/n))

            # SİNE-BLEND: UTS'ye pürüzsüz yaklaşım (Türev = 0)
            m_arr = (stressPlastic[1:] - stressPlastic[:-1]) / (strainPlastic[1:] - strainPlastic[:-1])
            found_idx = -1
            for i in range(len(strainPlastic)-2, 10, -1):
                L_actual = strainPlastic[-1] - strainPlastic[i]
                L_req = (np.pi / 2) * (UTS - stressPlastic[i]) / m_arr[i-1] if m_arr[i-1] > 0 else 0
                if L_actual >= L_req:
                    found_idx = i
                    break
            
            if found_idx != -1:
                s_start = strainPlastic[found_idx]
                str_start = stressPlastic[found_idx]
                L_sine = strainPlastic[-1] - s_start
                s_seg = strainPlastic[found_idx:] - s_start
                stressPlastic[found_idx:] = str_start + (UTS - str_start) * np.sin((np.pi / 2) * (s_seg / L_sine))

            self.stressUptoUTS = np.concatenate((stressElastic, stressDrop, stressPlateau, stressPlastic))
            self.strainUptoUTS = np.concatenate((strainElastic, strainDrop, strainPlateau, strainPlastic))
        else:
            sStart = 0.70 * yieldStrength
            
            stressUptoUTS = np.linspace(0, UTS, 2500)
            strainPMax = (UTS / K)**(1/n)
            deltaSMax = UTS - sStart
            if deltaSMax > 0:
                A = strainPMax / (deltaSMax ** (1/n))
            else:
                A = 0.0

            deltaS = np.maximum(0, stressUptoUTS - sStart)
            strainPlasticSmooth = A * (deltaS ** (1/n))
            
            self.stressUptoUTS = stressUptoUTS
            self.strainUptoUTS = (stressUptoUTS / E) + strainPlasticSmooth

            # SİNE-BLEND: UTS'ye pürüzsüz yaklaşım (Türev = 0)
            m_arr = (self.stressUptoUTS[1:] - self.stressUptoUTS[:-1]) / (self.strainUptoUTS[1:] - self.strainUptoUTS[:-1])
            found_idx = -1
            for i in range(len(self.strainUptoUTS)-2, 100, -1):
                L_actual = self.strainUptoUTS[-1] - self.strainUptoUTS[i]
                L_req = (np.pi / 2) * (UTS - self.stressUptoUTS[i]) / m_arr[i-1] if m_arr[i-1] > 0 else 0
                if L_actual >= L_req:
                    found_idx = i
                    break
            
            if found_idx != -1:
                s_start = self.strainUptoUTS[found_idx]
                str_start = self.stressUptoUTS[found_idx]
                L_sine = self.strainUptoUTS[-1] - s_start
                s_seg = self.strainUptoUTS[found_idx:] - s_start
                self.stressUptoUTS[found_idx:] = str_start + (UTS - str_start) * np.sin((np.pi / 2) * (s_seg / L_sine))

        self.strainatUTS = self.strainUptoUTS[-1]

        # --- DİNAMİK BOYUN VERME (NECKING) MODELİ ---
        self.strainNecking = np.linspace(self.strainatUTS, fractureStrain, 2000)[1:]
        deltaStrain = self.strainNecking - self.strainatUTS
        deltaStrainF = fractureStrain - self.strainatUTS
        
        if deltaStrainF > 0:
            # Malzemenin sünekliğine göre boyun verme eğrisinin derecesini (p) belirliyoruz:
            # - Eğer malzeme daha gevreksi (deltaStrainF küçükse -> Titanyum): p = 2 (Dik/Ani düşüş)
            # - Eğer malzeme daha sünekse (deltaStrainF büyükse -> Alüminyum/Çelik): p = 3.5 - 4 (Geniş plato + yumuşak düşüş)
            p = 2.0 + min(2.0, max(0.0, (deltaStrainF - 0.03) * 20))
            
            a = (fractureStress - UTS) / (deltaStrainF ** p)
            self.stressNecking = a * (deltaStrain ** p) + UTS
        else:
            self.stressNecking = np.array([fractureStress] * len(self.strainNecking))
       
        self.combinedStrain = np.concatenate((self.strainUptoUTS, self.strainNecking))
        self.combinedStress = np.concatenate((self.stressUptoUTS, self.stressNecking))

        self.combinedForce = self.combinedStress * self.defCrossSectionalArea
        self.combinedExtension = self.combinedStrain * defLength

        self.yieldIndex = (np.abs(self.stressUptoUTS - yieldStrength)).argmin()

        return (self.combinedForce.tolist(), self.combinedExtension.tolist(), 
                self.combinedStrain.tolist(), self.combinedStress.tolist(), 
                self.defCrossSectionalArea, int(self.yieldIndex))

    def runSimulator(self):
        if hasattr(self, 'comparisonData'):
            del self.comparisonData
        
        if self.dataFrameOfMaterials is None:
            messagebox.showerror("Database Error", "Please load an Excel database first using 'Load External Excel'!")
            return

        try:
            for attrName, attrVar, _, _ in self.parameters:
                entryWidget = getattr(self, attrName)
                val = float(entryWidget.get())
                setattr(self, attrVar, val)

            self.lowCarbonCheck = self.isLowCarbonSteel.get()

            self.forceList, self.extensionList, self.strainList, self.stressList, self.crossSectionalArea, self.placeOfYieldPoint = self.testEngine(
                self.diameter, self.length, self.E, self.yieldStrength, self.UTS, self.K, self.n, self.fractureStress, self.fractureStrain, self.lowCarbonCheck
            )

            self.toughnessOfMaterial = 0.0
            for i in range(0, len(self.strainList)-1):
                self.toughnessOfMaterial += ((self.stressList[i] + self.stressList[i+1]) / 2) * (self.strainList[i+1] - self.strainList[i]) 
            
            self.placeOfMaxStress = self.stressList.index(max(self.stressList))
            maxForcekN = max(self.forceList) / 1000
            maxDisplacementmm = self.extensionList[-1]
            elasticModulusGPa = self.E / 1000
            uniformElong = self.strainList[self.placeOfMaxStress] * 100
            elongatFracture = self.strainList[-1] * 100
            elongAfterFracture = (self.strainList[-1] - (self.stressList[-1] / self.E)) * 100
            modulusOfResilience = (self.yieldStrength ** 2) / (2 * self.E)

            self.selectedMaterial = self.localVariableOfMaterial.get()
            self.safeMaterialName = self.selectedMaterial.replace("/", "_").replace("\\", "_")

            self.resultFolder = os.path.join(BASE_DIR, "Tensile Tests", f"Results of {self.safeMaterialName}")
            os.makedirs(self.resultFolder, exist_ok=True)
            self.fileName = os.path.join(self.resultFolder, f"Tensile Test Results of {self.safeMaterialName}.csv")
            
            with open(self.fileName, "w", encoding="utf-8") as file:
                file.write("Force(N),Extension(mm),Stress(MPa),Strain(mm/mm)\n")
                for f, e, s, st in zip(self.forceList, self.extensionList, self.stressList, self.strainList):
                    file.write(f"{f:.4f},{e:.4f},{s:.4f},{st:.6f}\n")

            for item in self.resultTree.get_children():
                self.resultTree.delete(item)
                
            cat1 = self.resultTree.insert("", "end", text="Test Information", open=True)
            self.resultTree.insert(cat1, "end", text="Maximum Force", values=(f"{maxForcekN:.3f}", "kN"))
            self.resultTree.insert(cat1, "end", text="Maximum Displacement", values=(f"{maxDisplacementmm:.3f}", "mm"))
            
            cat2 = self.resultTree.insert("", "end", text="Mechanical Properties", open=True)
            self.resultTree.insert(cat2, "end", text="Yield Strength", values=(f"{self.yieldStrength:.2f}", "MPa"))
            self.resultTree.insert(cat2, "end", text="Ultimate Tensile Strength", values=(f"{self.UTS:.2f}", "MPa"))
            self.resultTree.insert(cat2, "end", text="Fracture Strength", values=(f"{self.fractureStress:.2f}", "MPa"))
            self.resultTree.insert(cat2, "end", text="Elastic Modulus", values=(f"{elasticModulusGPa:.2f}", "GPa"))
            self.resultTree.insert(cat2, "end", text="Uniform Elongation", values=(f"{uniformElong:.2f}", "%"))
            self.resultTree.insert(cat2, "end", text="Elongation at Fracture", values=(f"{elongatFracture:.2f}", "%"))
            self.resultTree.insert(cat2, "end", text="Elongation after Fracture", values=(f"{elongAfterFracture:.2f}", "%"))
            
            cat3 = self.resultTree.insert("", "end", text="Energy Properties", open=True)
            self.resultTree.insert(cat3, "end", text="Modulus of Toughness", values=(f"{self.toughnessOfMaterial:.2f}", "MJ/m³"))
            self.resultTree.insert(cat3, "end", text="Modulus of Resilience", values=(f"{modulusOfResilience:.3f}", "MJ/m³"))
            
            cat4 = self.resultTree.insert("", "end", text="Model Parameters", open=True)
            self.resultTree.insert(cat4, "end", text="K (Strength Coefficient)", values=(f"{self.K:.1f}", "MPa"))
            self.resultTree.insert(cat4, "end", text="n (Hardening Exponent)", values=(f"{self.n:.3f}", "-"))

            tableRows = []
            for catID in self.resultTree.get_children():
                catText = self.resultTree.item(catID, "text")
                for itemID in self.resultTree.get_children(catID):
                    itemValues = self.resultTree.item(itemID)
                    propName = itemValues["text"]
                    val, unit = itemValues["values"]
                    tableRows.append({"Category": catText, "Property": propName, "Value": val, "Unit": unit})
            
            summary_df = pd.DataFrame(tableRows)
            summaryFileName = os.path.join(self.resultFolder, f"Summary Report of {self.safeMaterialName}.csv")
            summary_df.to_csv(summaryFileName, index=False, encoding="utf-8")

            self.animIndex = 0
            self.startLiveAnimation()

        except ValueError:
            messagebox.showerror("Entry Error", "Please just enter a number!")
        except FileNotFoundError:
            messagebox.showerror("File Error", "tensilematerials.xlsx is not found!")
        except KeyError as e:
            messagebox.showerror("Key Error", f"This column was not found in Excel: {e}")
        except Exception as e:
            messagebox.showerror("System Error", str(e))

    def deleteSelectedMaterial(self):
        if self.currentFilePath is None or self.dataFrameOfMaterials is None:
            messagebox.showerror("Error", "No loaded database was found.")
            return
        self.selectedMat = self.localVariableOfMaterial.get()
        if self.selectedMat == "No Database Loaded" or not self.selectedMat:
            messagebox.showerror("Error", "No valid material selected for deletion!")
            return
        
        if not messagebox.askyesno("Confirmation", f"Are you sure about deleting this material: '{self.selectedMat}'?"):
            return
        try:
            self.dataFrameOfMaterials.drop(self.selectedMat, inplace=True)
            self.dataFrameOfMaterials.to_excel(self.currentFilePath)
            self.materialNames = self.dataFrameOfMaterials.index.tolist()

            menu = self.menuOfMaterial["menu"]
            menu.delete(0, "end")
            
            if self.materialNames:
                for name in self.materialNames:
                    menu.add_command(label=name, command=tk._setit(self.localVariableOfMaterial, name))
                self.localVariableOfMaterial.set(self.materialNames[0])
            else:
                self.materialNames = ["No Database Loaded"]
                self.localVariableOfMaterial.set("No Database Loaded")
                menu.add_command(label="No Database Loaded", command=tk._setit(self.localVariableOfMaterial, "No Database Loaded"))

            messagebox.showinfo("Success", f"'{self.selectedMat}' is deleted successfully!")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred in deleting: {e}")

    def saveNewMaterialToExcel(self):
        if self.currentFilePath is None:
            messagebox.showerror("Error", "Please first of all load a database via 'Load External Excel'")
            return
        self.newMatName = self.newMaterialNameEntry.get().strip()
        if not self.newMatName:
            messagebox.showerror("Error", "Please enter a name for the new material")
            return
        try:
            newData = {
                "moduleofyoung": float(self.youngEntry.get()),
                "yieldstrength": float(self.yieldStrengthEntry.get()),
                "ultimatetensilestrength": float(self.UTSEntry.get()),
                "strengthcoefficient": float(self.strengthCoefficientEntry.get()),
                "strainhardeningexponent": float(self.strainHardeningExponentEntry.get()),
                "fracturestress": float(self.fractureStressEntry.get()),
                "fracturestrain": float(self.fractureStrainEntry.get())
            }
            self.dataFrameOfMaterials.loc[self.newMatName] = newData
            self.dataFrameOfMaterials.to_excel(self.currentFilePath)

            self.materialNames = self.dataFrameOfMaterials.index.tolist()
            menu = self.menuOfMaterial["menu"]
            menu.delete(0, "end")
            for name in self.materialNames:
                menu.add_command(label=name, command=tk._setit(self.localVariableOfMaterial, name))

            self.localVariableOfMaterial.set(self.newMatName)
            messagebox.showinfo("Success", f"'{self.newMatName}' added to the Excel database successfully!")
        except ValueError:
            messagebox.showerror("Entry Error", "Please ensure that you enter valid numerical values into all attribute boxes!")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred during save: {e}")

    def updateEntries(self, *args):
        self.selectedMat = self.localVariableOfMaterial.get()
        if self.dataFrameOfMaterials is None or self.selectedMat == "No Database Loaded":
            return

        try:
            self.mat = self.dataFrameOfMaterials.loc[self.selectedMat]
            mappings = [
                (self.youngEntry, "moduleofyoung"),
                (self.yieldStrengthEntry, "yieldstrength"),
                (self.UTSEntry, "ultimatetensilestrength"),
                (self.strengthCoefficientEntry, "strengthcoefficient"),
                (self.strainHardeningExponentEntry, "strainhardeningexponent"),
                (self.fractureStressEntry, "fracturestress"),
                (self.fractureStrainEntry, "fracturestrain")
            ]
            for entryWidget, keyName in mappings:
                entryWidget.delete(0, tk.END)
                entryWidget.insert(0, str(self.mat[keyName]))
        
        except Exception as e:
            messagebox.showerror("Error", f"Failed to update entries: {e}")
                
    def loadExternalExcel(self):
        self.filePath = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if self.filePath:
            try:
                self.currentFilePath = self.filePath
                self.df = pd.read_excel(self.filePath)
                self.df.columns = [col.strip().lower().replace(" ", "").replace("-", "").replace("_", "").replace("/", "") for col in self.df.columns]
                
                if "materialname" in self.df.columns:
                    self.df.set_index("materialname", inplace=True)
                else:
                    raise KeyError("Excel dosyasında 'materialname' adında bir sütun bulunamadı!")
                
                self.dataFrameOfMaterials = self.df
                self.materialNames = self.dataFrameOfMaterials.index.tolist()

                menu = self.menuOfMaterial["menu"]
                menu.delete(0, "end")
                for name in self.materialNames:
                    menu.add_command(label=name, command=tk._setit(self.localVariableOfMaterial, name))

                self.localVariableOfMaterial.set(self.materialNames[0])
                messagebox.showinfo("Success!", "Successfully loaded the new material database!")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load Excel file: {e}")

    def saveGraphImage(self):
        if not hasattr(self, 'fig') or not hasattr(self, 'safeMaterialName') or not hasattr(self, 'resultFolder'):
            messagebox.showerror("Error", "No active graph found to save!")
            return
            
        self.graphFileName = os.path.join(self.resultFolder, f"Graph of {self.safeMaterialName}.png")
        self.fig.savefig(self.graphFileName, dpi=300, bbox_inches='tight')
        messagebox.showinfo("Saved", f"Graph successfully saved to:\n{self.graphFileName}")

    def compareAndShowResults(self):
        folder1 = filedialog.askdirectory(title="Select Folder for 1. Material")
        if not folder1:
            return
            
        folder2 = filedialog.askdirectory(title="Select Folder for 2. Material")
        if not folder2:
            return

        try:
            csv_files1 = [os.path.join(folder1, f) for f in os.listdir(folder1) if f.endswith('.csv') and not f.startswith("Summary Report")]
            csv_files2 = [os.path.join(folder2, f) for f in os.listdir(folder2) if f.endswith('.csv') and not f.startswith("Summary Report")]

            summary_files1 = [os.path.join(folder1, f) for f in os.listdir(folder1) if f.startswith("Summary Report")]
            summary_files2 = [os.path.join(folder2, f) for f in os.listdir(folder2) if f.startswith("Summary Report")]

            if not csv_files1 or not csv_files2 or not summary_files1 or not summary_files2:
                messagebox.showerror("Error", "Required raw data or Summary Report CSV files not found in one or both folders!")
                return

            df1 = pd.read_csv(csv_files1[0])
            df2 = pd.read_csv(csv_files2[0])
            
            name1 = os.path.basename(folder1).replace("Results of ", "")
            name2 = os.path.basename(folder2).replace("Results of ", "")

            self.comparisonData = {
                "file1": (df1["Extension(mm)"], df1["Force(N)"], df1["Strain(mm/mm)"], df1["Stress(MPa)"], name1),
                "file2": (df2["Extension(mm)"], df2["Force(N)"], df2["Strain(mm/mm)"], df2["Stress(MPa)"], name2)
            }

            sum_df1 = pd.read_csv(summary_files1[0])
            sum_df2 = pd.read_csv(summary_files2[0])

            for item in self.resultTree.get_children():
                self.resultTree.delete(item)

            cat_mat1 = self.resultTree.insert("", "end", text=f"📊 {name1}", open=True)
            cat_mat2 = self.resultTree.insert("", "end", text=f"📊 {name2}", open=True)

            for row1, row2 in zip(sum_df1.itertuples(), sum_df2.itertuples()):
                prop = row1.Property
                val1_str, unit1 = str(row1.Value), str(row1.Unit)
                val2_str, unit2 = str(row2.Value), str(row2.Unit)

                try:
                    num1 = float(val1_str)
                    num2 = float(val2_str)

                    if num1 > num2:
                        tag1, tag2 = "better", "worse"
                    elif num2 > num1:
                        tag1, tag2 = "worse", "better"
                    else:
                        tag1, tag2 = "equal", "equal"
                except ValueError:
                    tag1, tag2 = "equal", "equal"

                self.resultTree.insert(cat_mat1, "end", text=prop, values=(val1_str, unit1), tags=(tag1,))
                self.resultTree.insert(cat_mat2, "end", text=prop, values=(val2_str, unit2), tags=(tag2,))

            messagebox.showinfo("Success", "Comparison results loaded with color-coded analysis!")

        except Exception as e:
            messagebox.showerror("Error", f"Comparison failed: {e}")

    def showComparisonGraph(self):
        if not hasattr(self, 'comparisonData'):
            messagebox.showwarning("Warning", "Please compare CSV files first!")
            return 0
            
        self.tensilePage.pack_forget()
        self.graphPage.pack(fill=tk.BOTH, expand=True)

        for widget in self.graphCanvasFrame.winfo_children():
            widget.destroy()

        self.fig = Figure(figsize=(10, 5), dpi=100)
        self.comparisonAx1 = self.fig.add_subplot(121) 
        self.comparisonAx2 = self.fig.add_subplot(122) 

        self.comparisonAx1.set_title("Force-Extension Overlay (Live)")
        self.comparisonAx1.set_xlabel("Extension (mm)")
        self.comparisonAx1.set_ylabel("Force (N)")
        self.comparisonAx1.grid(True)

        self.comparisonAx2.set_title("Stress-Strain Overlay (Live)")
        self.comparisonAx2.set_xlabel("Strain (mm/mm)")
        self.comparisonAx2.set_ylabel("Stress (MPa)")
        self.comparisonAx2.grid(True)

        self.comparisonCanvas = FigureCanvasTkAgg(self.fig, master=self.graphCanvasFrame)
        self.comparisonCanvas.draw()
        
        self.comparisonToolbar = NavigationToolbar2Tk(self.comparisonCanvas, self.graphCanvasFrame)
        self.comparisonToolbar.update()
        
        self.comparisonCanvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

        self.comparisonAnimIndex = 0
        self.updateComparisonAnimation()

    def updateComparisonAnimation(self):
        ext1, force1, strain1, stress1, name1 = self.comparisonData["file1"]
        ext2, force2, strain2, stress2, name2 = self.comparisonData["file2"]
        
        ext1, force1, strain1, stress1 = np.array(ext1), np.array(force1), np.array(strain1), np.array(stress1)
        ext2, force2, strain2, stress2 = np.array(ext2), np.array(force2), np.array(strain2), np.array(stress2)

        totalSteps1 = len(ext1)
        totalSteps2 = len(ext2)
        maxSteps = max(totalSteps1, totalSteps2)
        
        if self.comparisonAnimIndex < maxSteps:
            stepChunk = max(1, maxSteps // 150)
            currentEnd = min(self.comparisonAnimIndex + stepChunk, maxSteps)

            end1 = min(currentEnd, totalSteps1)
            end2 = min(currentEnd, totalSteps2)

            liveExtension1 = ext1[:end1]
            liveForce1 = force1[:end1]
            liveStrain1 = strain1[:end1]
            liveStress1 = stress1[:end1]

            liveExtension2 = ext2[:end2]
            liveForce2 = force2[:end2]
            liveStrain2 = strain2[:end2]
            liveStress2 = stress2[:end2]

            self.comparisonAx1.clear()
            self.comparisonAx2.clear()

            self.comparisonAx1.plot(liveExtension1, liveForce1, label=name1, color="blue", linewidth=2)
            self.comparisonAx2.plot(liveStrain1, liveStress1, label=name1, color="blue", linewidth=2)

            self.comparisonAx1.plot(liveExtension2, liveForce2, label=name2, color="orange", linewidth=2)
            self.comparisonAx2.plot(liveStrain2, liveStress2, label=name2, color="orange", linewidth=2)

            maxIdx1 = force1.argmax()
            if end1 > maxIdx1:
                self.comparisonAx1.plot(ext1[maxIdx1], force1[maxIdx1], 'bo', markersize=5)
                self.comparisonAx2.plot(strain1[maxIdx1], stress1[maxIdx1], 'bo', markersize=5)

            fractureIdx1 = totalSteps1 - 1
            if end1 >= fractureIdx1:
                self.comparisonAx1.plot(ext1[fractureIdx1], force1[fractureIdx1], 'bx', markersize=6)
                self.comparisonAx2.plot(strain1[fractureIdx1], stress1[fractureIdx1], 'bx', markersize=6)

            maxIdx2 = force2.argmax()
            if end2 > maxIdx2:
                self.comparisonAx1.plot(ext2[maxIdx2], force2[maxIdx2], 'ro', markersize=5)
                self.comparisonAx2.plot(strain2[maxIdx2], stress2[maxIdx2], 'ro', markersize=5)

            fractureIdx2 = totalSteps2 - 1
            if end2 >= fractureIdx2:
                self.comparisonAx1.plot(ext2[fractureIdx2], force2[fractureIdx2], 'rx', markersize=6)
                self.comparisonAx2.plot(strain2[fractureIdx2], stress2[fractureIdx2], 'rx', markersize=6)

            self.comparisonAx1.set_title("Force-Extension Overlay (Live)")
            self.comparisonAx1.set_xlabel("Extension (mm)")
            self.comparisonAx1.set_ylabel("Force (N)")
            self.comparisonAx1.grid(True)
            self.comparisonAx1.legend()

            self.comparisonAx2.set_title("Stress-Strain Overlay (Live)")
            self.comparisonAx2.set_xlabel("Strain (mm/mm)")
            self.comparisonAx2.set_ylabel("Stress (MPa)")
            self.comparisonAx2.grid(True)
            self.comparisonAx2.legend()

            self.comparisonCanvas.draw()

            self.comparisonAnimIndex = currentEnd
            self.hub.mainPage.after(5, self.updateComparisonAnimation)


class ElectricalAnalyzer:
    def __init__(self, referenceOfHub):
        self.currentFilePath = None       
        self.hub = referenceOfHub
       
        self.dataFrameOfMaterials = None
        self.materialNames = ["No Database Loaded"]

        self.localVariableOfMaterial = tk.StringVar(self.hub.container)
        self.localVariableOfMaterial.set(self.materialNames[0])

        self.setupElectricalPage()
        self.setupGraphPage()  
        self.checkAndCreateDatabase()

    def checkAndCreateDatabase(self):
        electricalFolder = "Electrical Tests"
        os.makedirs(os.path.join(BASE_DIR, electricalFolder), exist_ok=True)
        
        self.currentFilePath = os.path.join(BASE_DIR, electricalFolder, "electricalmaterials.xlsx")
        
        if not os.path.exists(self.currentFilePath):
            try:
                defaultData = {
                    "materialname": [
                        "Gold (Au) Thin Film",
                        "Copper (Cu) Thin Film",
                        "Aluminum (Al) Thin Film",
                        "P-Type Silicon Wafer",
                        "N-Type Silicon Wafer",
                        "GaAs (Gallium Arsenide) Wafer",
                        "ITO (Indium Tin Oxide)",
                        "FTO (Fluorine-doped Tin Oxide)"
                    ],
                    "bulkresistivity": [0.00000244, 0.00000168, 0.00000265, 1.5, 5.0, 0.01, 0.0001, 0.0005],
                    "bandgap": [0, 0, 0, 1.12, 1.12, 1.42, 3.6, 3.25]
                }
                defaultDF = pd.DataFrame(defaultData)
                defaultDF.set_index("materialname", inplace=True)
                defaultDF.to_excel(self.currentFilePath)
            except Exception as e:
                print(f"Could not create default electrical database: {e}")

        try:
            self.dataFrameOfMaterials = pd.read_excel(self.currentFilePath)
            self.dataFrameOfMaterials.columns = [
                col.strip().lower().replace(" ", "").replace("-", "").replace("_", "").replace("/", "") 
                for col in self.dataFrameOfMaterials.columns
            ]
            if "materialname" in self.dataFrameOfMaterials.columns:
                self.dataFrameOfMaterials.set_index("materialname", inplace=True)

            self.materialNames = self.dataFrameOfMaterials.index.tolist()
            menu = self.menuOfMaterial["menu"]
            menu.delete(0, "end")
            for name in self.materialNames:
                menu.add_command(label=name, command=tk._setit(self.localVariableOfMaterial, name))
            self.localVariableOfMaterial.set(self.materialNames[0])
        except Exception as e:
            print(f"Could not load electrical database: {e}")

    def showHubPage(self):
        self.electricalPage.pack_forget()
        self.graphPage.pack_forget()
        self.hub.hubPage.pack(fill=tk.BOTH, expand=True)

    def showElectricalPage(self):
        self.hub.hubPage.pack_forget()
        self.graphPage.pack_forget()
        self.electricalPage.pack(fill=tk.BOTH, expand=True)

    def setupElectricalPage(self):
        self.electricalPage = tk.Frame(self.hub.container, bg="#2b2b2b")
        
        topNav = tk.Frame(self.electricalPage, bg="#808080", height=45)
        topNav.pack(fill=tk.X, side=tk.TOP)

        btnBackToHub = tk.Button(topNav, text="⬅ Back to Hub", font=("Arial", 10, "bold"), bg="white", fg="black", command=self.showHubPage)
        btnBackToHub.pack(side=tk.LEFT, padx=15, pady=6)

        tk.Label(topNav, text="4-Point Probe & I-V Characterization Module", font=("Arial", 12, "bold"), bg="#808080", fg="white").pack(side=tk.LEFT, padx=15)

        contentFrame = tk.Frame(self.electricalPage, bg="#2b2b2b")
        contentFrame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        self.leftControlFrame = tk.Frame(contentFrame, bg="#2b2b2b")
        self.leftControlFrame.pack(side=tk.LEFT, fill=tk.Y, padx=15, pady=10)

        self.mainLabel = tk.Label(self.leftControlFrame, text="Choose a Material:", font=("Arial", 14, "bold"), bg="#2b2b2b", fg="white")
        self.mainLabel.grid(row=0, column=0, columnspan=2, pady=(5,5))

        self.localVariableOfMaterial = tk.StringVar(self.leftControlFrame)
        self.localVariableOfMaterial.set(list(self.materialNames)[0])
        self.menuOfMaterial = tk.OptionMenu(self.leftControlFrame, self.localVariableOfMaterial, *self.materialNames)
        self.menuOfMaterial.config(width=18, bg="white", fg="black", activebackground="white", activeforeground="black")
        self.menuOfMaterial.grid(row=1, column=0, columnspan=2, pady=(0,10))

        self.parameters = [
            ("bulkResistivityEntry", "bulkResistivity", "Bulk Resistivity (Ohm-cm):", 2),
            ("bandgapEntry", "bandgap", "Bandgap Energy (eV):", 3), 
            ("filmThicknessEntry", "filmThickness", "Film Thickness (t) (nm):", 4), 
            ("maxVoltageEntry", "sweepMaxVoltage", "Sweep Max Voltage (V):", 5),
            ("currentStepEntry", "measurementSteps", "Measurement Steps:", 6),
            ("probeSpacingEntry", "probeSpacing", "Probe Spacing (s) (mm):", 7),
            ("temperatureEntry", "temperature", "Temperature (°C):", 8)
        ]

        for attrName, attrVar, labelText, rowIdx in self.parameters:
            tk.Label(self.leftControlFrame, text=labelText, font=("Arial", 9, "bold"), bg="#2b2b2b", fg="white").grid(row=rowIdx, column=0, sticky="w", padx=5, pady=3)
            entry = tk.Entry(self.leftControlFrame, width=12, bg="white", fg="black", insertbackground="black")
            entry.grid(row=rowIdx, column=1, padx=5, pady=3)
            setattr(self, attrName, entry)

        self.filmThicknessEntry.insert(0, "100.0")
        self.maxVoltageEntry.insert(0, "5.0")
        self.currentStepEntry.insert(0, "100")
        self.probeSpacingEntry.insert(0, "1.0")
        self.temperatureEntry.insert(0, "25.0")

        self.isSemiConductor = tk.BooleanVar(value=False)
        self.semiConductorCheckButton = tk.Checkbutton(
            self.leftControlFrame, text="Semiconductor (Non-Linear I-V)", 
            variable=self.isSemiConductor, font=("Arial", 9, "bold"), 
            fg="#ff6b6b", bg="#2b2b2b", activebackground="#2b2b2b", selectcolor="white"
        )
        self.semiConductorCheckButton.grid(row=9, column=0, columnspan=2, pady=5)

        self.runButton = tk.Button(self.leftControlFrame, text="-Run Simulation-", bg="green", fg="black", font=("Arial", 14, "bold"), command=self.runSimulator, width=40)
        self.runButton.grid(row=10, column=0, columnspan=2, pady=8)

        self.uploadButton = tk.Button(self.leftControlFrame, text="Load Excel", command=self.loadExternalExcel, bg="blue", fg="black", font=("Arial", 9, "bold"), width=12)
        self.uploadButton.grid(row=11, column=0, padx=2, pady=3)

        self.deleteButton = tk.Button(self.leftControlFrame, text="Delete Selected", command=self.deleteSelectedMaterial, bg="red", fg="black", font=("Arial", 9, "bold"), width=18)
        self.deleteButton.grid(row=11, column=1, padx=2, pady=3)

        tk.Label(self.leftControlFrame, text="New Material Name:", font=("Arial", 9, "bold"), fg="#87ceeb", bg="#2b2b2b").grid(row=12, column=0, sticky="w", padx=5, pady=(8, 2))
        self.newMaterialNameEntry = tk.Entry(self.leftControlFrame, width=12, bg="white", fg="black", insertbackground="black")
        self.newMaterialNameEntry.grid(row=12, column=1, padx=5, pady=(8, 2))

        self.saveButton = tk.Button(self.leftControlFrame, text="Save the New Material", bg="orange", fg="black", font=("Arial", 9, "bold"), command=self.saveNewMaterialToExcel, width=20)
        self.saveButton.grid(row=13, column=0, columnspan=2, padx=2, pady=5)

        self.compareButton = tk.Button(self.leftControlFrame, text="Compare CSVs", command=self.compareAndShowResults, bg="purple", fg="black", font=("Arial", 9, "bold"), width=30)
        self.compareButton.grid(row=14, column=0, columnspan=2, pady=3)

        self.localVariableOfMaterial.trace_add("write", self.updateEntries)

        self.rightResultFrame = tk.Frame(contentFrame, relief="sunken", bd=1, bg="#2b2b2b")
        self.rightResultFrame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=15, pady=10)

        tk.Label(self.rightResultFrame, text="Analysis Results", font=("Arial", 14, "bold"), fg="#fbfbfb", bg="#2b2b2b").pack(pady=10)
        
        self.resultTree = ttk.Treeview(self.rightResultFrame, columns=("Value", "Unit"), selectmode="none", height=23)
        self.resultTree.heading("#0", text="Property")
        self.resultTree.heading("Value", text="Value")
        self.resultTree.heading("Unit", text="Unit")
        self.resultTree.column("#0", width=250, anchor="w")
        self.resultTree.column("Value", width=110, anchor="center")
        self.resultTree.column("Unit", width=80, anchor="center")
        
        self.resultTree.tag_configure("better", foreground="green")
        self.resultTree.tag_configure("worse", foreground="red")
        self.resultTree.tag_configure("equal", foreground="orange")
        
        self.resultTree.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        self.btnGraphForward = tk.Button(self.rightResultFrame, text="Show Graph 📈 ➔", font=("Arial", 11, "bold"), bg="white", fg="black", command=self.showGraphPage)
        self.btnGraphForward.pack(anchor="e", padx=10, pady=10)

    def setupGraphPage(self):
        self.graphPage = tk.Frame(self.hub.container, bg="#2b2b2b")
        
        graphTopNav = tk.Frame(self.graphPage, bg="white", height=45)
        graphTopNav.pack(fill=tk.X, side=tk.TOP)

        btnBackToResults = tk.Button(graphTopNav, text="Back to Results", font=("Arial", 10, "bold"), bg="gray", fg="black", command=self.showElectricalPage)
        btnBackToResults.pack(side=tk.LEFT, padx=15, pady=6)

        saveGraphBtn = tk.Button(graphTopNav, text="Save Graph Image", font=("Arial", 10, "bold"), bg="orange", fg="black", command=self.saveGraphImage)
        saveGraphBtn.pack(side=tk.RIGHT, padx=15, pady=6)

        self.graphCanvasFrame = tk.Frame(self.graphPage, bg="white")
        self.graphCanvasFrame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

    def showGraphPage(self):
        if hasattr(self, 'comparisonData') and self.comparisonData:
            self.showComparisonGraph()
            return

        if not hasattr(self, 'fig'):
            messagebox.showwarning("Warning", "Please run a simulation or compare CSV files first!")
            return

        self.electricalPage.pack_forget()
        self.graphPage.pack(fill=tk.BOTH, expand=True)

        for widget in self.graphCanvasFrame.winfo_children():
            widget.destroy()

        canvas = FigureCanvasTkAgg(self.fig, master=self.graphCanvasFrame)
        canvas.draw()

        toolbar = NavigationToolbar2Tk(canvas, self.graphCanvasFrame)
        toolbar.update()

        canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

    def updateEntries(self, *args):
        self.selectedMat = self.localVariableOfMaterial.get()
        if self.dataFrameOfMaterials is None or self.selectedMat == "No Database Loaded":
            return

        try:
            self.mat = self.dataFrameOfMaterials.loc[self.selectedMat]
            mappings = [
                (self.bulkResistivityEntry, "bulkresistivity"),
                (self.bandgapEntry, "bandgap")
            ]
            for entryWidget, keyName in mappings:
                entryWidget.delete(0, tk.END)
                entryWidget.insert(0, str(self.mat[keyName]))
        
        except Exception as e:
            messagebox.showerror("Error", f"Failed to update entries: {e}")

    def loadExternalExcel(self):
        self.filePath = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if self.filePath:
            try:
                self.currentFilePath = self.filePath
                self.df = pd.read_excel(self.filePath)
                self.df.columns = [col.strip().lower().replace(" ", "").replace("-", "").replace("_", "").replace("/", "") for col in self.df.columns]

                if "materialname" in self.df.columns:
                    self.df.set_index("materialname", inplace=True)
                else:
                    raise KeyError("Excel dosyasında 'materialname' adında bir sütun bulunamadı!")

                self.dataFrameOfMaterials = self.df
                self.materialNames = self.dataFrameOfMaterials.index.tolist()

                menu = self.menuOfMaterial["menu"]
                menu.delete(0, "end")
                for name in self.materialNames:
                    menu.add_command(label=name, command=tk._setit(self.localVariableOfMaterial, name))

                self.localVariableOfMaterial.set(self.materialNames[0])
                messagebox.showinfo("Success!", "Successfully loaded the new material database!")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load Excel file: {e}")

    def deleteSelectedMaterial(self):
        if self.currentFilePath is None or self.dataFrameOfMaterials is None:
            messagebox.showerror("Error", "No loaded database was found.")
            return
        self.selectedMat = self.localVariableOfMaterial.get()
        if self.selectedMat == "No Database Loaded" or not self.selectedMat:
            messagebox.showerror("Error", "No valid material selected for deletion!")
            return
        
        if not messagebox.askyesno("Confirmation", f"Are you sure about deleting this material: '{self.selectedMat}'?"):
            return
        try:
            self.dataFrameOfMaterials.drop(self.selectedMat, inplace=True)
            self.dataFrameOfMaterials.to_excel(self.currentFilePath)
            self.materialNames = self.dataFrameOfMaterials.index.tolist()

            menu = self.menuOfMaterial["menu"]
            menu.delete(0, "end")
            
            if self.materialNames:
                for name in self.materialNames:
                    menu.add_command(label=name, command=tk._setit(self.localVariableOfMaterial, name))
                self.localVariableOfMaterial.set(self.materialNames[0])
            else:
                self.materialNames = ["No Database Loaded"]
                self.localVariableOfMaterial.set("No Database Loaded")
                menu.add_command(label="No Database Loaded", command=tk._setit(self.localVariableOfMaterial, "No Database Loaded"))

            messagebox.showinfo("Success", f"'{self.selectedMat}' is deleted successfully!")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred in deleting: {e}")

    def saveNewMaterialToExcel(self):
        if self.currentFilePath is None:
            messagebox.showerror("Error", "Please first of all load a database via 'Load External Excel'")
            return
        self.newMatName = self.newMaterialNameEntry.get().strip()
        if not self.newMatName:
            messagebox.showerror("Error", "Please enter a name for the new material")
            return
        try:
            newData = {
                "bulkresistivity": float(self.bulkResistivityEntry.get()),
                "bandgap": float(self.bandgapEntry.get()),
            }
            self.dataFrameOfMaterials.loc[self.newMatName] = newData
            self.dataFrameOfMaterials.to_excel(self.currentFilePath)

            self.materialNames = self.dataFrameOfMaterials.index.tolist()
            menu = self.menuOfMaterial["menu"]
            menu.delete(0, "end")
            for name in self.materialNames:
                menu.add_command(label=name, command=tk._setit(self.localVariableOfMaterial, name))

            self.localVariableOfMaterial.set(self.newMatName)
            messagebox.showinfo("Success", f"'{self.newMatName}' added to the Excel database successfully!")
        except ValueError:
            messagebox.showerror("Entry Error", "Please ensure that you enter valid numerical values into all attribute boxes!")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred during save: {e}")

    def testEngine(self, defBulkResistivity, defFilmThickness, defProbeSpacing, defTemperature):
        defFilmThickness = defFilmThickness * 1.0e-7  # nm -> cm
        defProbeSpacing = defProbeSpacing * 0.1       # mm -> cm

        sheetResistance = defBulkResistivity / defFilmThickness
        theoriticalResistance = (sheetResistance * math.log(2)) / math.pi

        self.arrayofVoltage = np.linspace(0, self.sweepMaxVoltage, int(self.measurementSteps))

        if not self.isSemiConductor.get():
            self.arrayofCurrent = self.arrayofVoltage / theoriticalResistance
        else:
            q = 1.602e-19
            kB = 1.3806e-23
            n = 1.5

            defTemperature = defTemperature + 273.15
            self.thermalVoltage = (kB * defTemperature) / q

            scaleFactor = 1e-3
            reverseSaturationCurrent = scaleFactor * np.exp(-(q * self.bandgap) / (2 * kB * defTemperature))
            
            Rs = 10.0  # Ohm
            raw_current = reverseSaturationCurrent * (np.exp(np.minimum(self.arrayofVoltage / (n * self.thermalVoltage), 40)) - 1)
            
            compliance_current = 0.1  # Amper
            self.arrayofCurrent = np.minimum(raw_current / (1 + raw_current * Rs / np.maximum(self.arrayofVoltage, 1e-6)), compliance_current)

        return (
            self.arrayofVoltage.tolist(), 
            self.arrayofCurrent.tolist(), 
            sheetResistance, 
            theoriticalResistance
        )

    def runSimulator(self):
        if hasattr(self, 'comparisonData'):
            del self.comparisonData

        if self.dataFrameOfMaterials is None:
            messagebox.showerror("Database Error", "Please load an Excel database first using 'Load External Excel'!")
            return
        
        try:
            for attrName, attrVar, _, _ in self.parameters:
                entryWidget = getattr(self, attrName)
                val = float(entryWidget.get())
                if attrVar == "measurementSteps":
                    val = int(val)
                setattr(self, attrVar, val)

            self.listofVoltage, self.listofCurrent, self.sheetResistance, self.theoriticalResistance = self.testEngine(
                self.bulkResistivity, self.filmThickness, self.probeSpacing, self.temperature
            )

            self.selectedMaterial = self.localVariableOfMaterial.get()
            self.safeMaterialName = self.selectedMaterial.replace("/", "_").replace("\\", "_")
            self.resultFolder = os.path.join(BASE_DIR, "Electrical Tests", f"Results of {self.safeMaterialName}")
            os.makedirs(self.resultFolder, exist_ok=True)

            # Save Raw Data CSV for Compare functionality
            self.fileName = os.path.join(self.resultFolder, f"Electrical Test Results of {self.safeMaterialName}.csv")
            with open(self.fileName, "w", encoding="utf-8") as file:
                file.write("Voltage(V),Current(A),Current(mA)\n")
                for v, i in zip(self.listofVoltage, self.listofCurrent):
                    file.write(f"{v:.4f},{i:.8e},{i*1000.0:.6f}\n")

            for item in self.resultTree.get_children():
                self.resultTree.delete(item)

            maxCurrent_mA = max(self.listofCurrent) * 1000 if self.listofCurrent else 0
            conductance_S = (1.0 / self.theoriticalResistance) if self.theoriticalResistance > 0 else 0.0

            cat1 = self.resultTree.insert("", "end", text="Test & Geometry Information", open=True)
            self.resultTree.insert(cat1, "end", text="Film Thickness (t)", values=(f"{self.filmThickness:.2f}", "nm"))
            self.resultTree.insert(cat1, "end", text="Probe Spacing (s)", values=(f"{self.probeSpacing:.2f}", "mm"))
            self.resultTree.insert(cat1, "end", text="Temperature", values=(f"{self.temperature:.1f}", "°C"))

            cat2 = self.resultTree.insert("", "end", text="Electrical Properties", open=True)
            self.resultTree.insert(cat2, "end", text="Bulk Resistivity (ρ)", values=(f"{self.bulkResistivity:.4e}", "Ω·cm"))
            self.resultTree.insert(cat2, "end", text="Sheet Resistance (Rs)", values=(f"{self.sheetResistance:.2f}", "Ω/sq"))
            self.resultTree.insert(cat2, "end", text="Calculated Resistance (R)", values=(f"{self.theoriticalResistance:.2f}", "Ω"))
            self.resultTree.insert(cat2, "end", text="Electrical Conductance (G)", values=(f"{conductance_S:.4e}", "S"))

            cat3 = self.resultTree.insert("", "end", text="Measurement Parameters", open=True)
            mode_text = "Semiconductor (Non-Linear)" if self.isSemiConductor.get() else "Ohmic (Linear)"
            self.resultTree.insert(cat3, "end", text="Measurement Mode", values=(mode_text, "-"))
            self.resultTree.insert(cat3, "end", text="Max Applied Voltage", values=(f"{self.sweepMaxVoltage:.2f}", "V"))
            self.resultTree.insert(cat3, "end", text="Max Measured Current", values=(f"{maxCurrent_mA:.4f}", "mA"))

            if self.isSemiConductor.get():
                thermalVoltage_mV = getattr(self, 'thermalVoltage', 0) * 1000
                self.resultTree.insert(cat3, "end", text="Thermal Voltage (VT)", values=(f"{thermalVoltage_mV:.2f}", "mV"))
                self.resultTree.insert(cat3, "end", text="Bandgap Energy (Eg)", values=(f"{self.bandgap:.2f}", "eV"))

            tableRows = []
            for catID in self.resultTree.get_children():
                catText = self.resultTree.item(catID, "text")
                for itemID in self.resultTree.get_children(catID):
                    itemValues = self.resultTree.item(itemID)
                    propName = itemValues["text"]
                    val, unit = itemValues["values"]
                    tableRows.append({"Category": catText, "Property": propName, "Value": val, "Unit": unit})
            
            summary_df = pd.DataFrame(tableRows)
            summaryFileName = os.path.join(self.resultFolder, f"Summary Report of {self.safeMaterialName}.csv")
            summary_df.to_csv(summaryFileName, index=False, encoding="utf-8")

            self.animIndex = 0
            self.startLiveAnimation()

        except ValueError:
            messagebox.showerror("Entry Error", "Please enter valid numeric values for all input fields!")
        except Exception as e:
            messagebox.showerror("Simulation Error", f"An error occurred during simulation: {e}")

    def compareAndShowResults(self):
        folder1 = filedialog.askdirectory(title="Select Folder for 1. Material")
        if not folder1:
            return
            
        folder2 = filedialog.askdirectory(title="Select Folder for 2. Material")
        if not folder2:
            return

        try:
            csv_files1 = [os.path.join(folder1, f) for f in os.listdir(folder1) if f.endswith('.csv') and not f.startswith("Summary Report")]
            csv_files2 = [os.path.join(folder2, f) for f in os.listdir(folder2) if f.endswith('.csv') and not f.startswith("Summary Report")]

            summary_files1 = [os.path.join(folder1, f) for f in os.listdir(folder1) if f.startswith("Summary Report")]
            summary_files2 = [os.path.join(folder2, f) for f in os.listdir(folder2) if f.startswith("Summary Report")]

            if not csv_files1 or not csv_files2 or not summary_files1 or not summary_files2:
                messagebox.showerror("Error", "Required raw data or Summary Report CSV files not found in one or both folders!")
                return

            df1 = pd.read_csv(csv_files1[0])
            df2 = pd.read_csv(csv_files2[0])
            
            name1 = os.path.basename(folder1).replace("Results of ", "")
            name2 = os.path.basename(folder2).replace("Results of ", "")

            self.comparisonData = {
                "file1": (df1["Voltage(V)"], df1["Current(mA)"], name1),
                "file2": (df2["Voltage(V)"], df2["Current(mA)"], name2)
            }

            sum_df1 = pd.read_csv(summary_files1[0])
            sum_df2 = pd.read_csv(summary_files2[0])

            for item in self.resultTree.get_children():
                self.resultTree.delete(item)

            cat_mat1 = self.resultTree.insert("", "end", text=f"📊 {name1}", open=True)
            cat_mat2 = self.resultTree.insert("", "end", text=f"📊 {name2}", open=True)

            # Properties where LOWER value is better (e.g. Resistivity, Resistance)
            lower_is_better = ["Bulk Resistivity (ρ)", "Sheet Resistance (Rs)", "Calculated Resistance (R)"]

            for row1, row2 in zip(sum_df1.itertuples(), sum_df2.itertuples()):
                prop = row1.Property
                val1_str, unit1 = str(row1.Value), str(row1.Unit)
                val2_str, unit2 = str(row2.Value), str(row2.Unit)

                try:
                    num1 = float(val1_str)
                    num2 = float(val2_str)

                    if prop in lower_is_better:
                        if num1 < num2:
                            tag1, tag2 = "better", "worse"
                        elif num2 < num1:
                            tag1, tag2 = "worse", "better"
                        else:
                            tag1, tag2 = "equal", "equal"
                    else:
                        if num1 > num2:
                            tag1, tag2 = "better", "worse"
                        elif num2 > num1:
                            tag1, tag2 = "worse", "better"
                        else:
                            tag1, tag2 = "equal", "equal"
                except ValueError:
                    tag1, tag2 = "equal", "equal"

                self.resultTree.insert(cat_mat1, "end", text=prop, values=(val1_str, unit1), tags=(tag1,))
                self.resultTree.insert(cat_mat2, "end", text=prop, values=(val2_str, unit2), tags=(tag2,))

            messagebox.showinfo("Success", "Electrical comparison results loaded with color-coded analysis!")

        except Exception as e:
            messagebox.showerror("Error", f"Comparison failed: {e}")

    def showComparisonGraph(self):
        if not hasattr(self, 'comparisonData'):
            messagebox.showwarning("Warning", "Please compare CSV files first!")
            return
            
        self.electricalPage.pack_forget()
        self.graphPage.pack(fill=tk.BOTH, expand=True)

        for widget in self.graphCanvasFrame.winfo_children():
            widget.destroy()

        self.fig = Figure(figsize=(10, 5), dpi=100)
        self.comparisonAx1 = self.fig.add_subplot(121) 
        self.comparisonAx2 = self.fig.add_subplot(122) 

        self.comparisonCanvas = FigureCanvasTkAgg(self.fig, master=self.graphCanvasFrame)
        self.comparisonCanvas.draw()
        
        self.comparisonToolbar = NavigationToolbar2Tk(self.comparisonCanvas, self.graphCanvasFrame)
        self.comparisonToolbar.update()
        
        self.comparisonCanvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

        self.comparisonAnimIndex = 0
        self.updateComparisonAnimation()

    def updateComparisonAnimation(self):
        v1, i1_mA, name1 = self.comparisonData["file1"]
        v2, i2_mA, name2 = self.comparisonData["file2"]
        
        v1, i1_mA = np.array(v1), np.array(i1_mA)
        v2, i2_mA = np.array(v2), np.array(i2_mA)

        totalSteps1 = len(v1)
        totalSteps2 = len(v2)
        maxSteps = max(totalSteps1, totalSteps2)

        if self.comparisonAnimIndex < maxSteps:
            stepChunk = max(1, maxSteps // 150)
            currentEnd = min(self.comparisonAnimIndex + stepChunk, maxSteps)

            end1 = min(currentEnd, totalSteps1)
            end2 = min(currentEnd, totalSteps2)

            self.comparisonAx1.clear()
            self.comparisonAx2.clear()

            self.comparisonAx1.plot(v1[:end1], i1_mA[:end1], label=name1, color="blue", linewidth=2)
            self.comparisonAx1.plot(v2[:end2], i2_mA[:end2], label=name2, color="orange", linewidth=2)

            self.comparisonAx1.set_title("I-V Characterization Overlay (Live)")
            self.comparisonAx1.set_xlabel("Voltage (V)")
            self.comparisonAx1.set_ylabel("Current (mA)")
            self.comparisonAx1.grid(True)
            self.comparisonAx1.legend()

            self.comparisonAx2.semilogy(v1[:end1], np.maximum(i1_mA[:end1], 1e-9), label=name1, color="blue", linewidth=2)
            self.comparisonAx2.semilogy(v2[:end2], np.maximum(i2_mA[:end2], 1e-9), label=name2, color="orange", linewidth=2)

            self.comparisonAx2.set_title("Semi-Log I-V Overlay (Log Scale)")
            self.comparisonAx2.set_xlabel("Voltage (V)")
            self.comparisonAx2.set_ylabel("Current (mA)")
            self.comparisonAx2.grid(True, which="both", ls="--")
            self.comparisonAx2.legend()

            self.comparisonCanvas.draw()

            self.comparisonAnimIndex = currentEnd
            self.hub.mainPage.after(5, self.updateComparisonAnimation)

    def saveGraphImage(self):
        if not hasattr(self, 'fig') or not hasattr(self, 'safeMaterialName') or not hasattr(self, 'resultFolder'):
            messagebox.showerror("Error", "No active graph found to save! Please run simulation first.")
            return
            
        try:
            self.graphFileName = os.path.join(self.resultFolder, f"Graph of {self.safeMaterialName}.png")
            self.fig.savefig(self.graphFileName, dpi=300, bbox_inches='tight')
            messagebox.showinfo("Saved", f"Graph successfully saved to:\n{self.graphFileName}")
        except Exception as e:
            messagebox.showerror("Save Error", f"Could not save graph image: {e}")

    def startLiveAnimation(self):
        self.electricalPage.pack_forget()
        self.graphPage.pack(fill=tk.BOTH, expand=True)
        if hasattr(self, 'fig'):
            self.fig.clf() # Eski figürün içeriğini temizle

        for widget in self.graphCanvasFrame.winfo_children():
            widget.destroy()

        self.fig = Figure(figsize=(10, 5), dpi=100)

        if not self.isSemiConductor.get():
            self.ax1 = self.fig.add_subplot(111)
        else:
            self.ax1 = self.fig.add_subplot(121)
            self.ax2 = self.fig.add_subplot(122)

        self.canvas = FigureCanvasTkAgg(self.fig, master=self.graphCanvasFrame)
        self.canvas.draw()
        
        self.toolbar = NavigationToolbar2Tk(self.canvas, self.graphCanvasFrame)
        self.toolbar.update()
        
        self.canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

        self.updateAnimation()

    def updateAnimation(self):
        totalSteps = len(self.listofVoltage)

        if self.animIndex < totalSteps:
            stepChunk = 1
            currentEnd = min(self.animIndex + stepChunk, totalSteps)

            voltageSlice = self.listofVoltage[:currentEnd]
            currentSlice = self.listofCurrent[:currentEnd]

            currentSlicemA = np.array(currentSlice) * 1000
            currentSlicemA = np.maximum(currentSlicemA, 1e-9)

            if not self.isSemiConductor.get():
                self.ax1.clear()
                self.ax1.plot(voltageSlice, currentSlicemA, color="red", linewidth=2)
                self.ax1.set_title("Graph of I-V (Live)")
                self.ax1.set_xlabel("Voltage (V)")
                self.ax1.set_ylabel("Current (mA)")
                self.ax1.grid(True)
            else:
                self.ax1.clear()
                self.ax2.clear()

                self.ax1.plot(voltageSlice, currentSlicemA, color="red", linewidth=2)
                self.ax1.set_title("Graph of Voltage-Current (Live)")
                self.ax1.set_xlabel("Voltage (V)")
                self.ax1.set_ylabel("Current (mA)")
                self.ax1.grid(True)

                self.ax2.semilogy(voltageSlice, currentSlicemA, color="blue", linewidth=2)
                self.ax2.set_title("Semi-Log I-V Graph (Log Scale)")
                self.ax2.set_xlabel("Voltage (V)")
                self.ax2.set_ylabel("Current (mA)")
                self.ax2.grid(True, which="both", ls="--")

            self.canvas.draw()
            self.animIndex = currentEnd

            if self.animIndex < totalSteps:
                self.hub.mainPage.after(5, self.updateAnimation)
            else:
                self.vMax = self.listofVoltage[-1]
                self.iMaxmA = self.listofCurrent[-1] * 1000

                self.ax1.plot(self.vMax, self.iMaxmA, 'go', markersize=6)
                self.ax1.annotate(f"Maximum Current: {self.iMaxmA:.3f} mA", xy=(self.vMax, self.iMaxmA), 
                                  xytext=(-120, -20), textcoords="offset points", 
                                  color="green", fontweight="bold")

                if self.isSemiConductor.get():
                    self.ax2.plot(self.vMax, max(self.iMaxmA, 1e-9), 'go', markersize=6)

                self.canvas.draw()
                messagebox.showinfo("Success", "Voltage sweep simulation completed successfully!")


class XRDAnalyzer:
    def __init__(self, referenceOfHub):
        self.hub = referenceOfHub

        self.currentFilePath = None       
        self.dataFrameOfMaterials = None
        self.materialNames = ["No Database Loaded"]

        self.localVariableOfMaterial = tk.StringVar(self.hub.container)
        self.localVariableOfMaterial.set(self.materialNames[0])

        self.setupXRDPage()
        self.setupGraphPage()  
        self.checkAndCreateDatabase()

    def checkAndCreateDatabase(self):
        xrdFolder = "XRD Tests"
        os.makedirs(os.path.join(BASE_DIR, xrdFolder), exist_ok=True)
        
        self.currentFilePath = os.path.join(BASE_DIR, xrdFolder, "xrdmaterials.xlsx")
        
        if not os.path.exists(self.currentFilePath):
            try:
                defaultData = {
                    "materialname": [
                        "Copper (Cu) [FCC]",
                        "Aluminum (Al) [FCC]",
                        "Iron (Alpha-Fe) [BCC]",
                        "Titanium (Alpha-Ti) [HCP]",
                        "Tin (Beta-Sn) [Tetragonal]",
                        "Gallium (Ga) [Orthorhombic]",
                        "Polonium (Po) [SC]"
                    ],
                    "crystalsystem": ["Cubic", "Cubic", "Cubic", "Hexagonal", "Tetragonal", "Orthorhombic", "Cubic"],
                    "bravaistype": ["FCC", "FCC", "BCC", "HCP", "I (Body-Centered)", "C (Base-Centered)", "SC"],
                    "latticeparametera": [0.3615, 0.4049, 0.2866, 0.2950, 0.5832, 0.4519, 0.3359],
                    "latticeparameterb": [0.0, 0.0, 0.0, 0.0, 0.0, 0.4526, 0.0],
                    "latticeparameterc": [0.0, 0.0, 0.0, 0.4686, 0.3181, 0.7657, 0.0]
                }
                defaultDF = pd.DataFrame(defaultData)
                defaultDF.set_index("materialname", inplace=True)
                defaultDF.to_excel(self.currentFilePath)
            except Exception as e:
                print(f"Could not create default XRD database: {e}")

        try:
            self.dataFrameOfMaterials = pd.read_excel(self.currentFilePath)
            self.dataFrameOfMaterials.columns = [
                col.strip().lower().replace(" ", "").replace("-", "").replace("_", "").replace("/", "") 
                for col in self.dataFrameOfMaterials.columns
            ]
            if "materialname" in self.dataFrameOfMaterials.columns:
                self.dataFrameOfMaterials.set_index("materialname", inplace=True)

            self.materialNames = self.dataFrameOfMaterials.index.tolist()
            menu = self.menuOfMaterial["menu"]
            menu.delete(0, "end")
            for name in self.materialNames:
                menu.add_command(label=name, command=tk._setit(self.localVariableOfMaterial, name))
            self.localVariableOfMaterial.set(self.materialNames[0])
        except Exception as e:
            print(f"Could not load XRD database: {e}")

    def showHubPage(self):
        self.xrdPage.pack_forget()
        self.graphPage.pack_forget()
        self.hub.hubPage.pack(fill=tk.BOTH, expand=True)

    def showXRDPage(self):
        self.hub.hubPage.pack_forget()
        self.graphPage.pack_forget()
        self.xrdPage.pack(fill=tk.BOTH, expand=True)

    def setupXRDPage(self):
        self.xrdPage = tk.Frame(self.hub.container, bg="#2b2b2b")
        
        topNav = tk.Frame(self.xrdPage, bg="#808080", height=45)
        topNav.pack(fill=tk.X, side=tk.TOP)

        btnBackToHub = tk.Button(topNav, text="⬅ Back to Hub", font=("Arial", 10, "bold"), bg="white", fg="black", command=self.showHubPage)
        btnBackToHub.pack(side=tk.LEFT, padx=15, pady=6)

        tk.Label(topNav, text="X-Ray Diffraction (XRD) Crystallography Module", font=("Arial", 12, "bold"), bg="#808080", fg="white").pack(side=tk.LEFT, padx=15)

        contentFrame = tk.Frame(self.xrdPage, bg="#2b2b2b")
        contentFrame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        self.leftControlFrame = tk.Frame(contentFrame, bg="#2b2b2b")
        self.leftControlFrame.pack(side=tk.LEFT, fill=tk.Y, padx=15, pady=10)

        self.mainLabel = tk.Label(self.leftControlFrame, text="Choose a Material:", font=("Arial", 14, "bold"), bg="#2b2b2b", fg="white")
        self.mainLabel.grid(row=0, column=0, columnspan=2, pady=(5,5))

        self.localVariableOfMaterial = tk.StringVar(self.leftControlFrame)
        self.localVariableOfMaterial.set(list(self.materialNames)[0])
        self.menuOfMaterial = tk.OptionMenu(self.leftControlFrame, self.localVariableOfMaterial, *self.materialNames)
        self.menuOfMaterial.config(width=18, bg="white", fg="black", activebackground="white", activeforeground="black")
        self.menuOfMaterial.grid(row=1, column=0, columnspan=2, pady=(0,10))

        # Kristal Sistemi Seçimi
        tk.Label(self.leftControlFrame, text="Crystal System:", font=("Arial", 9, "bold"), bg="#2b2b2b", fg="white").grid(row=2, column=0, sticky="w", padx=5, pady=3)
        self.crystalSystemVar = tk.StringVar(value="Cubic")
        self.crystalSystemMenu = ttk.Combobox(self.leftControlFrame, textvariable=self.crystalSystemVar, values=["Cubic", "Tetragonal", "Orthorhombic", "Hexagonal"], state="readonly", width=12)
        self.crystalSystemMenu.grid(row=2, column=1, padx=5, pady=3)
        self.crystalSystemMenu.bind("<<ComboboxSelected>>", self.onCrystalSystemChange)

        # Bravais Örgü / Yapı Tipi
        tk.Label(self.leftControlFrame, text="Crystal Structure:", font=("Arial", 9, "bold"), bg="#2b2b2b", fg="white").grid(row=3, column=0, sticky="w", padx=5, pady=3)
        self.crystalStructureVar = tk.StringVar(value="FCC")
        self.crystalStructureMenu = ttk.Combobox(self.leftControlFrame, textvariable=self.crystalStructureVar, values=["SC", "BCC", "FCC"], state="readonly", width=12)
        self.crystalStructureMenu.grid(row=3, column=1, padx=5, pady=3)

        # Örgü Parametreleri (a, b, c)
        tk.Label(self.leftControlFrame, text="Lattice Parameter a (nm):", font=("Arial", 9, "bold"), bg="#2b2b2b", fg="white").grid(row=4, column=0, sticky="w", padx=5, pady=3)
        self.latticeParameterAEntry = tk.Entry(self.leftControlFrame, width=12, bg="white", fg="black", insertbackground="black")
        self.latticeParameterAEntry.grid(row=4, column=1, padx=5, pady=3)

        tk.Label(self.leftControlFrame, text="Lattice Parameter b (nm):", font=("Arial", 9, "bold"), bg="#2b2b2b", fg="white").grid(row=5, column=0, sticky="w", padx=5, pady=3)
        self.latticeParameterBEntry = tk.Entry(self.leftControlFrame, width=12, bg="white", fg="black", insertbackground="black")
        self.latticeParameterBEntry.grid(row=5, column=1, padx=5, pady=3)

        tk.Label(self.leftControlFrame, text="Lattice Parameter c (nm):", font=("Arial", 9, "bold"), bg="#2b2b2b", fg="white").grid(row=6, column=0, sticky="w", padx=5, pady=3)
        self.latticeParameterCEntry = tk.Entry(self.leftControlFrame, width=12, bg="white", fg="black", insertbackground="black")
        self.latticeParameterCEntry.grid(row=6, column=1, padx=5, pady=3)

        self.parameters = [
            ("wavelengthEntry", "wavelength", "Wavelength λ (nm):", 7),
            ("minTwoThetaEntry", "minTwoTheta", "Min 2θ (°):", 8),
            ("maxTwoThetaEntry", "maxTwoTheta", "Max 2θ (°):", 9),
            ("fwhmEntry", "fwhm", "Peak FWHM (°):", 10)
        ]

        for attrName, attrVar, labelText, rowIdx in self.parameters:
            tk.Label(self.leftControlFrame, text=labelText, font=("Arial", 9, "bold"), bg="#2b2b2b", fg="white").grid(row=rowIdx, column=0, sticky="w", padx=5, pady=3)
            entry = tk.Entry(self.leftControlFrame, width=12, bg="white", fg="black", insertbackground="black")
            entry.grid(row=rowIdx, column=1, padx=5, pady=3)
            setattr(self, attrName, entry)

        self.wavelengthEntry.insert(0, "0.15418")
        self.minTwoThetaEntry.insert(0, "20.0")
        self.maxTwoThetaEntry.insert(0, "90.0")
        self.fwhmEntry.insert(0, "0.2")

        self.runButton = tk.Button(self.leftControlFrame, text="-Run Simulation-", bg="green", fg="black", font=("Arial", 14, "bold"), command=self.runSimulator, width=40)
        self.runButton.grid(row=11, column=0, columnspan=2, pady=12)

        self.uploadButton = tk.Button(self.leftControlFrame, text="Load Excel", command=self.loadExternalExcel, bg="blue", fg="black", font=("Arial", 9, "bold"), width=12)
        self.uploadButton.grid(row=12, column=0, padx=2, pady=3)

        self.deleteButton = tk.Button(self.leftControlFrame, text="Delete Selected", command=self.deleteSelectedMaterial, bg="red", fg="black", font=("Arial", 9, "bold"), width=18)
        self.deleteButton.grid(row=12, column=1, padx=2, pady=3)

        tk.Label(self.leftControlFrame, text="New Material Name:", font=("Arial", 9, "bold"), fg="#87ceeb", bg="#2b2b2b").grid(row=13, column=0, sticky="w", padx=5, pady=(8, 2))
        self.newMaterialNameEntry = tk.Entry(self.leftControlFrame, width=12, bg="white", fg="black", insertbackground="black")
        self.newMaterialNameEntry.grid(row=13, column=1, padx=5, pady=(8, 2))

        self.saveButton = tk.Button(self.leftControlFrame, text="Save the New Material", bg="orange", fg="black", font=("Arial", 9, "bold"), command=self.saveNewMaterialToExcel, width=20)
        self.saveButton.grid(row=14, column=0, columnspan=2, padx=2, pady=5)

        self.compareButton = tk.Button(self.leftControlFrame, text="Compare CSVs", command=self.compareAndShowResults, bg="purple", fg="black", font=("Arial", 9, "bold"), width=30)
        self.compareButton.grid(row=15, column=0, columnspan=2, pady=3)

        self.localVariableOfMaterial.trace_add("write", self.updateEntries)

        self.rightResultFrame = tk.Frame(contentFrame, relief="sunken", bd=1, bg="#2b2b2b")
        self.rightResultFrame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=15, pady=10)

        tk.Label(self.rightResultFrame, text="Analysis Results", font=("Arial", 14, "bold"), fg="#fbfbfb", bg="#2b2b2b").pack(pady=10)
        
        self.resultTree = ttk.Treeview(self.rightResultFrame, columns=("Value", "Unit"), selectmode="none", height=23)
        self.resultTree.heading("#0", text="Property / Peak")
        self.resultTree.heading("Value", text="Value")
        self.resultTree.heading("Unit", text="Unit")
        self.resultTree.column("#0", width=250, anchor="w")
        self.resultTree.column("Value", width=110, anchor="center")
        self.resultTree.column("Unit", width=80, anchor="center")
        
        self.resultTree.tag_configure("better", foreground="green")
        self.resultTree.tag_configure("worse", foreground="red")
        self.resultTree.tag_configure("equal", foreground="orange")
        
        self.resultTree.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        self.btnGraphForward = tk.Button(self.rightResultFrame, text="Show Graph 📈 ➔", font=("Arial", 11, "bold"), bg="white", fg="black", command=self.showGraphPage)
        self.btnGraphForward.pack(anchor="e", padx=10, pady=10)

        self.onCrystalSystemChange()

    def onCrystalSystemChange(self, event=None):
        system = self.crystalSystemVar.get()
        
        if system == "Cubic":
            self.crystalStructureMenu['values'] = ["SC", "BCC", "FCC"]
            if self.crystalStructureVar.get() not in ["SC", "BCC", "FCC"]:
                self.crystalStructureVar.set("FCC")
            
            self.latticeParameterAEntry.config(state="normal")
            self.latticeParameterBEntry.delete(0, tk.END)
            self.latticeParameterBEntry.insert(0, "0.0")
            self.latticeParameterBEntry.config(state="disabled")
            
            self.latticeParameterCEntry.delete(0, tk.END)
            self.latticeParameterCEntry.insert(0, "0.0")
            self.latticeParameterCEntry.config(state="disabled")

        elif system == "Tetragonal":
            self.crystalStructureMenu['values'] = ["P (Simple)", "I (Body-Centered)"]
            if self.crystalStructureVar.get() not in ["P (Simple)", "I (Body-Centered)"]:
                self.crystalStructureVar.set("P (Simple)")

            self.latticeParameterAEntry.config(state="normal")
            self.latticeParameterBEntry.delete(0, tk.END)
            self.latticeParameterBEntry.insert(0, "0.0")
            self.latticeParameterBEntry.config(state="disabled")
            
            self.latticeParameterCEntry.config(state="normal")

        elif system == "Orthorhombic":
            self.crystalStructureMenu['values'] = ["P (Simple)", "I (Body-Centered)", "F (Face-Centered)", "C (Base-Centered)"]
            if self.crystalStructureVar.get() not in ["P (Simple)", "I (Body-Centered)", "F (Face-Centered)", "C (Base-Centered)"]:
                self.crystalStructureVar.set("P (Simple)")

            self.latticeParameterAEntry.config(state="normal")
            self.latticeParameterBEntry.config(state="normal")
            self.latticeParameterCEntry.config(state="normal")

        elif system == "Hexagonal":
            self.crystalStructureMenu['values'] = ["Simple Hexagonal", "HCP"]
            if self.crystalStructureVar.get() not in ["Simple Hexagonal", "HCP"]:
                self.crystalStructureVar.set("HCP")

            self.latticeParameterAEntry.config(state="normal")
            self.latticeParameterBEntry.delete(0, tk.END)
            self.latticeParameterBEntry.insert(0, "0.0")
            self.latticeParameterBEntry.config(state="disabled")
            
            self.latticeParameterCEntry.config(state="normal")

    def setupGraphPage(self):
        self.graphPage = tk.Frame(self.hub.container, bg="#2b2b2b")
        
        graphTopNav = tk.Frame(self.graphPage, bg="white", height=45)
        graphTopNav.pack(fill=tk.X, side=tk.TOP)

        btnBackToResults = tk.Button(graphTopNav, text="Back to Results", font=("Arial", 10, "bold"), bg="gray", fg="black", command=self.showXRDPage)
        btnBackToResults.pack(side=tk.LEFT, padx=15, pady=6)

        saveGraphBtn = tk.Button(graphTopNav, text="Save Graph Image", font=("Arial", 10, "bold"), bg="orange", fg="black", command=self.saveGraphImage)
        saveGraphBtn.pack(side=tk.RIGHT, padx=15, pady=6)

        self.graphCanvasFrame = tk.Frame(self.graphPage, bg="white")
        self.graphCanvasFrame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

    def showGraphPage(self):
        if hasattr(self, 'comparisonData') and self.comparisonData:
            self.showComparisonGraph()
            return

        if not hasattr(self, 'fig'):
            messagebox.showwarning("Warning", "Please run a simulation or compare CSV files first!")
            return

        self.xrdPage.pack_forget()
        self.graphPage.pack(fill=tk.BOTH, expand=True)

        for widget in self.graphCanvasFrame.winfo_children():
            widget.destroy()

        canvas = FigureCanvasTkAgg(self.fig, master=self.graphCanvasFrame)
        canvas.draw()

        toolbar = NavigationToolbar2Tk(canvas, self.graphCanvasFrame)
        toolbar.update()

        canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

    def updateEntries(self, *args):
        self.selectedMat = self.localVariableOfMaterial.get()
        if self.dataFrameOfMaterials is None or self.selectedMat == "No Database Loaded":
            return

        try:
            self.mat = self.dataFrameOfMaterials.loc[self.selectedMat]
            
            system = str(self.mat.get("crystalsystem", "Cubic")).capitalize()
            if system in ["Cubic", "Tetragonal", "Orthorhombic", "Hexagonal"]:
                self.crystalSystemVar.set(system)
            self.onCrystalSystemChange()

            if "bravaistype" in self.mat:
                self.crystalStructureVar.set(str(self.mat["bravaistype"]))

            a_val = self.mat.get("latticeparametera", self.mat.get("latticeparameter", 0.35))
            b_val = self.mat.get("latticeparameterb", 0.0)
            c_val = self.mat.get("latticeparameterc", 0.0)

            self.latticeParameterAEntry.config(state="normal")
            self.latticeParameterAEntry.delete(0, tk.END)
            self.latticeParameterAEntry.insert(0, str(a_val))

            if self.crystalSystemVar.get() == "Orthorhombic":
                self.latticeParameterBEntry.config(state="normal")
                self.latticeParameterBEntry.delete(0, tk.END)
                self.latticeParameterBEntry.insert(0, str(b_val))

            if self.crystalSystemVar.get() in ["Tetragonal", "Orthorhombic", "Hexagonal"]:
                self.latticeParameterCEntry.config(state="normal")
                self.latticeParameterCEntry.delete(0, tk.END)
                self.latticeParameterCEntry.insert(0, str(c_val))

        except Exception as e:
            messagebox.showerror("Error", f"Failed to update entries: {e}")

    def loadExternalExcel(self):
        self.filePath = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if self.filePath:
            try:
                self.currentFilePath = self.filePath
                self.df = pd.read_excel(self.filePath)
                self.df.columns = [col.strip().lower().replace(" ", "").replace("-", "").replace("_", "").replace("/", "") for col in self.df.columns]

                if "materialname" in self.df.columns:
                    self.df.set_index("materialname", inplace=True)
                else:
                    raise KeyError("Excel dosyasında 'materialname' adında bir sütun bulunamadı!")

                self.dataFrameOfMaterials = self.df
                self.materialNames = self.dataFrameOfMaterials.index.tolist()

                menu = self.menuOfMaterial["menu"]
                menu.delete(0, "end")
                for name in self.materialNames:
                    menu.add_command(label=name, command=tk._setit(self.localVariableOfMaterial, name))

                self.localVariableOfMaterial.set(self.materialNames[0])
                messagebox.showinfo("Success!", "Successfully loaded the new material database!")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load Excel file: {e}")

    def deleteSelectedMaterial(self):
        if self.currentFilePath is None or self.dataFrameOfMaterials is None:
            messagebox.showerror("Error", "No loaded database was found.")
            return
        self.selectedMat = self.localVariableOfMaterial.get()
        if self.selectedMat == "No Database Loaded" or not self.selectedMat:
            messagebox.showerror("Error", "No valid material selected for deletion!")
            return
        
        if not messagebox.askyesno("Confirmation", f"Are you sure about deleting this material: '{self.selectedMat}'?"):
            return
        try:
            self.dataFrameOfMaterials.drop(self.selectedMat, inplace=True)
            self.dataFrameOfMaterials.to_excel(self.currentFilePath)
            self.materialNames = self.dataFrameOfMaterials.index.tolist()

            menu = self.menuOfMaterial["menu"]
            menu.delete(0, "end")
            
            if self.materialNames:
                for name in self.materialNames:
                    menu.add_command(label=name, command=tk._setit(self.localVariableOfMaterial, name))
                self.localVariableOfMaterial.set(self.materialNames[0])
            else:
                self.materialNames = ["No Database Loaded"]
                self.localVariableOfMaterial.set("No Database Loaded")
                menu.add_command(label="No Database Loaded", command=tk._setit(self.localVariableOfMaterial, "No Database Loaded"))

            messagebox.showinfo("Success", f"'{self.selectedMat}' is deleted successfully!")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred in deleting: {e}")

    def saveNewMaterialToExcel(self):
        if self.currentFilePath is None:
            messagebox.showerror("Error", "Please first of all load a database via 'Load External Excel'")
            return
        self.newMatName = self.newMaterialNameEntry.get().strip()
        if not self.newMatName:
            messagebox.showerror("Error", "Please enter a name for the new material")
            return
        try:
            newData = {
                "crystalsystem": self.crystalSystemVar.get(),
                "bravaistype": self.crystalStructureVar.get(),
                "latticeparametera": float(self.latticeParameterAEntry.get()),
                "latticeparameterb": float(self.latticeParameterBEntry.get()),
                "latticeparameterc": float(self.latticeParameterCEntry.get())
            }
            self.dataFrameOfMaterials.loc[self.newMatName] = newData
            self.dataFrameOfMaterials.to_excel(self.currentFilePath)

            self.materialNames = self.dataFrameOfMaterials.index.tolist()
            menu = self.menuOfMaterial["menu"]
            menu.delete(0, "end")
            for name in self.materialNames:
                menu.add_command(label=name, command=tk._setit(self.localVariableOfMaterial, name))

            self.localVariableOfMaterial.set(self.newMatName)
            messagebox.showinfo("Success", f"'{self.newMatName}' added to the Excel database successfully!")
        except ValueError:
            messagebox.showerror("Entry Error", "Please ensure that you enter valid numerical values into all attribute boxes!")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred during save: {e}")

    def isValidReflection(self, h, k, l, crystalSystem, bravaisType):
        if h == 0 and k == 0 and l == 0:
            return False
            
        bravais = bravaisType.upper().split()[0]
        
        if crystalSystem == "Cubic":
            s = h**2 + k**2 + l**2
            if s in [7, 15, 23, 31, 39, 47, 55, 63]:
                return False
            if bravais in ["SC", "P"]:
                return True
            elif bravais in ["BCC", "I"]:
                return (h + k + l) % 2 == 0
            elif bravais in ["FCC", "F"]:
                h_even = (h % 2 == 0)
                k_even = (k % 2 == 0)
                l_even = (l % 2 == 0)
                return (h_even == k_even == l_even)
                
        elif crystalSystem == "Tetragonal":
            if bravais in ["P", "SC"]:
                return True
            elif bravais in ["I", "BCC"]:
                return (h + k + l) % 2 == 0
                
        elif crystalSystem == "Orthorhombic":
            if bravais == "P":
                return True
            elif bravais == "I":
                return (h + k + l) % 2 == 0
            elif bravais == "F":
                h_even = (h % 2 == 0)
                k_even = (k % 2 == 0)
                l_even = (l % 2 == 0)
                return (h_even == k_even == l_even)
            elif bravais == "C":
                return (h + k) % 2 == 0
                
        elif crystalSystem == "Hexagonal":
            if bravais == "HCP":
                if (h + 2*k) % 3 == 0 and (l % 2 != 0):
                    return False
                return True
            else:
                return True
                
        return True

    def calculateDSpacing(self, h, k, l, a, b, c, crystalSystem):
        if crystalSystem == "Cubic":
            return a / math.sqrt(h**2 + k**2 + l**2)
        elif crystalSystem == "Tetragonal":
            inv_d2 = (h**2 + k**2) / (a**2) + (l**2) / (c**2)
            return 1.0 / math.sqrt(inv_d2)
        elif crystalSystem == "Orthorhombic":
            inv_d2 = (h**2) / (a**2) + (k**2) / (b**2) + (l**2) / (c**2)
            return 1.0 / math.sqrt(inv_d2)
        elif crystalSystem == "Hexagonal":
            inv_d2 = (4.0 / 3.0) * (h**2 + h*k + k**2) / (a**2) + (l**2) / (c**2)
            return 1.0 / math.sqrt(inv_d2)
        return 0.0

    def getMultiplicity(self, h, k, l, crystalSystem):
        h, k, l = abs(h), abs(k), abs(l)
        
        if crystalSystem == "Cubic":
            vals = sorted([h, k, l])
            if vals == [0, 0, 0]: return 0
            if vals[0] == 0 and vals[1] == 0: return 6
            if vals[0] == 0 and vals[1] == vals[2]: return 12
            if vals[0] == 0 and vals[1] != vals[2]: return 24
            if vals[0] == vals[1] == vals[2]: return 8
            if vals[0] == vals[1] or vals[1] == vals[2]: return 24
            return 48
            
        elif crystalSystem == "Tetragonal":
            if h == 0 and k == 0 and l != 0: return 2
            if l == 0:
                if h == 0 or k == 0: return 4
                if h == k: return 4
                return 8
            else:
                if h == 0 or k == 0: return 8
                if h == k: return 8
                return 16
                
        elif crystalSystem == "Orthorhombic":
            zeros = [h, k, l].count(0)
            if zeros == 2: return 2
            if zeros == 1: return 4
            return 8
            
        elif crystalSystem == "Hexagonal":
            if h == 0 and k == 0 and l != 0: return 2
            if l == 0:
                if h == 0 or k == 0 or h == k: return 6
                return 12
            else:
                if h == 0 or k == 0 or h == k: return 12
                return 24
                
        return 1

    def testEngine(self, a, b, c, crystalSystem, bravaisType, wavelength, minTwoTheta, maxTwoTheta, fwhm):
        arrayOfTwoTheta = np.linspace(minTwoTheta, maxTwoTheta, 2500)
        arrayOfIntensity = np.zeros_like(arrayOfTwoTheta)
        detectedPeaks = []
        seen_d = set()

        h_range = range(-5, 6) if crystalSystem in ["Hexagonal", "Orthorhombic", "Tetragonal"] else range(0, 6)
        k_range = range(-5, 6) if crystalSystem in ["Hexagonal", "Orthorhombic", "Tetragonal"] else range(0, 6)
        l_range = range(0, 6)

        for h, k, l in product(h_range, k_range, l_range):
            if h == 0 and k == 0 and l == 0:
                continue

            if not self.isValidReflection(h, k, l, crystalSystem, bravaisType):
                continue

            d_hkl = self.calculateDSpacing(h, k, l, a, b, c, crystalSystem)
            if d_hkl <= 0:
                continue

            round_d = round(d_hkl, 5)
            if round_d in seen_d:
                continue

            sin_theta = wavelength / (2.0 * d_hkl)
            if sin_theta > 1.0 or sin_theta <= 0.0:
                continue

            theta_rad = math.asin(sin_theta)
            two_theta_deg = math.degrees(2.0 * theta_rad)

            if minTwoTheta <= two_theta_deg <= maxTwoTheta:
                seen_d.add(round_d)
                lp = (1.0 + math.cos(2.0 * theta_rad)**2) / (math.sin(theta_rad)**2 * math.cos(theta_rad))
                mult = self.getMultiplicity(h, k, l, crystalSystem)
                raw_intensity = mult * lp

                detectedPeaks.append((two_theta_deg, (h, k, l), d_hkl, raw_intensity))

        detectedPeaks.sort(key=lambda x: x[0])

        if detectedPeaks:
            max_raw = max(p[3] for p in detectedPeaks)
            detectedPeaks = [(p[0], p[1], p[2], (p[3] / max_raw) * 100.0) for p in detectedPeaks]

        sigma = fwhm / 2.355
        for two_theta_deg, hkl, d_hkl, rel_int in detectedPeaks:
            gaussian = rel_int * np.exp(-((arrayOfTwoTheta - two_theta_deg)**2) / (2.0 * sigma**2))
            arrayOfIntensity += gaussian

        if np.max(arrayOfIntensity) > 0:
            arrayOfIntensity = (arrayOfIntensity / np.max(arrayOfIntensity)) * 100.0

        return arrayOfTwoTheta.tolist(), arrayOfIntensity.tolist(), detectedPeaks

    def runSimulator(self):
        if hasattr(self, 'comparisonData'):
            del self.comparisonData

        if self.dataFrameOfMaterials is None:
            messagebox.showerror("Database Error", "Please load an Excel database first using 'Load External Excel'!")
            return
        
        try:
            self.crystalSystem = self.crystalSystemVar.get()
            self.crystalStructure = self.crystalStructureVar.get()
            self.latA = float(self.latticeParameterAEntry.get())
            self.latB = float(self.latticeParameterBEntry.get())
            self.latC = float(self.latticeParameterCEntry.get())
            self.wavelength = float(self.wavelengthEntry.get())
            self.minTwoTheta = float(self.minTwoThetaEntry.get())
            self.maxTwoTheta = float(self.maxTwoThetaEntry.get())
            self.fwhm = float(self.fwhmEntry.get())

            self.listofTwoTheta, self.listofIntensity, self.detectedPeaks = self.testEngine(
                self.latA, self.latB, self.latC, self.crystalSystem, self.crystalStructure, 
                self.wavelength, self.minTwoTheta, self.maxTwoTheta, self.fwhm
            )

            self.selectedMaterial = self.localVariableOfMaterial.get()
            self.safeMaterialName = self.selectedMaterial.replace("/", "_").replace("\\", "_")
            self.resultFolder = os.path.join(BASE_DIR, "XRD Tests", f"Results of {self.safeMaterialName}")
            os.makedirs(self.resultFolder, exist_ok=True)

            # Save Raw Data CSV for Compare functionality
            self.fileName = os.path.join(self.resultFolder, f"XRD Test Results of {self.safeMaterialName}.csv")
            with open(self.fileName, "w", encoding="utf-8") as file:
                file.write("TwoTheta(deg),Intensity(%)\n")
                for tt, intensity in zip(self.listofTwoTheta, self.listofIntensity):
                    file.write(f"{tt:.4f},{intensity:.4f}\n")

            for item in self.resultTree.get_children():
                self.resultTree.delete(item)

            cat1 = self.resultTree.insert("", "end", text="Crystal System Information", open=True)
            self.resultTree.insert(cat1, "end", text="Crystal System", values=(self.crystalSystem, "-"))
            self.resultTree.insert(cat1, "end", text="Crystal Structure", values=(self.crystalStructure, "-"))
            self.resultTree.insert(cat1, "end", text="Lattice Parameter a", values=(f"{self.latA:.4f}", "nm"))
            if self.crystalSystem == "Orthorhombic":
                self.resultTree.insert(cat1, "end", text="Lattice Parameter b", values=(f"{self.latB:.4f}", "nm"))
            if self.crystalSystem in ["Tetragonal", "Orthorhombic", "Hexagonal"]:
                self.resultTree.insert(cat1, "end", text="Lattice Parameter c", values=(f"{self.latC:.4f}", "nm"))
            self.resultTree.insert(cat1, "end", text="X-Ray Wavelength (λ)", values=(f"{self.wavelength:.5f}", "nm"))

            cat2 = self.resultTree.insert("", "end", text="Detected Diffraction Peaks", open=True)
            for two_theta_deg, hkl, d_hkl, rel_int in self.detectedPeaks:
                hkl_str = f"({hkl[0]} {hkl[1]} {hkl[2]})"
                self.resultTree.insert(cat2, "end", text=f"Peak {hkl_str}", values=(f"{two_theta_deg:.2f}° | d={d_hkl:.4f}", f"{rel_int:.1f}%"))

            summaryRows = [
                {"Category": "Crystal Info", "Property": "Crystal System", "Value": self.crystalSystem, "Unit": "-"},
                {"Category": "Crystal Info", "Property": "Crystal Structure", "Value": self.crystalStructure, "Unit": "-"},
                {"Category": "Crystal Info", "Property": "Lattice Parameter a", "Value": self.latA, "Unit": "nm"},
                {"Category": "Crystal Info", "Property": "Lattice Parameter b", "Value": self.latB, "Unit": "nm"},
                {"Category": "Crystal Info", "Property": "Lattice Parameter c", "Value": self.latC, "Unit": "nm"},
                {"Category": "Crystal Info", "Property": "Wavelength (λ)", "Value": self.wavelength, "Unit": "nm"}
            ]
            for two_theta_deg, hkl, d_hkl, rel_int in self.detectedPeaks:
                summaryRows.append({
                    "Category": "Diffraction Peak",
                    "Property": f"Peak ({hkl[0]} {hkl[1]} {hkl[2]})",
                    "Value": f"2θ={two_theta_deg:.2f}°, d={d_hkl:.4f}nm",
                    "Unit": f"I={rel_int:.1f}%"
                })

            summary_df = pd.DataFrame(summaryRows)
            summaryFileName = os.path.join(self.resultFolder, f"Summary Report of {self.safeMaterialName}.csv")
            summary_df.to_csv(summaryFileName, index=False, encoding="utf-8")

            self.animIndex = 0
            self.startLiveAnimation()

        except ValueError:
            messagebox.showerror("Entry Error", "Please enter valid numeric values for all input fields!")
        except Exception as e:
            messagebox.showerror("Simulation Error", f"An error occurred during simulation: {e}")

    def compareAndShowResults(self):
        folder1 = filedialog.askdirectory(title="Select Folder for 1. Material")
        if not folder1:
            return
            
        folder2 = filedialog.askdirectory(title="Select Folder for 2. Material")
        if not folder2:
            return

        try:
            csv_files1 = [os.path.join(folder1, f) for f in os.listdir(folder1) if f.endswith('.csv') and not f.startswith("Summary Report")]
            csv_files2 = [os.path.join(folder2, f) for f in os.listdir(folder2) if f.endswith('.csv') and not f.startswith("Summary Report")]

            summary_files1 = [os.path.join(folder1, f) for f in os.listdir(folder1) if f.startswith("Summary Report")]
            summary_files2 = [os.path.join(folder2, f) for f in os.listdir(folder2) if f.startswith("Summary Report")]

            if not csv_files1 or not csv_files2 or not summary_files1 or not summary_files2:
                messagebox.showerror("Error", "Required raw data or Summary Report CSV files not found in one or both folders!")
                return

            df1 = pd.read_csv(csv_files1[0])
            df2 = pd.read_csv(csv_files2[0])
            
            name1 = os.path.basename(folder1).replace("Results of ", "")
            name2 = os.path.basename(folder2).replace("Results of ", "")

            self.comparisonData = {
                "file1": (df1["TwoTheta(deg)"], df1["Intensity(%)"], name1),
                "file2": (df2["TwoTheta(deg)"], df2["Intensity(%)"], name2)
            }

            sum_df1 = pd.read_csv(summary_files1[0])
            sum_df2 = pd.read_csv(summary_files2[0])

            for item in self.resultTree.get_children():
                self.resultTree.delete(item)

            cat_mat1 = self.resultTree.insert("", "end", text=f"📊 {name1}", open=True)
            cat_mat2 = self.resultTree.insert("", "end", text=f"📊 {name2}", open=True)

            for row1, row2 in zip(sum_df1.itertuples(), sum_df2.itertuples()):
                prop = row1.Property
                val1_str, unit1 = str(row1.Value), str(row1.Unit)
                val2_str, unit2 = str(row2.Value), str(row2.Unit)

                try:
                    num1 = float(val1_str)
                    num2 = float(val2_str)

                    if num1 > num2:
                        tag1, tag2 = "better", "worse"
                    elif num2 > num1:
                        tag1, tag2 = "worse", "better"
                    else:
                        tag1, tag2 = "equal", "equal"
                except ValueError:
                    tag1, tag2 = "equal", "equal"

                self.resultTree.insert(cat_mat1, "end", text=prop, values=(val1_str, unit1), tags=(tag1,))
                self.resultTree.insert(cat_mat2, "end", text=prop, values=(val2_str, unit2), tags=(tag2,))

            messagebox.showinfo("Success", "XRD comparison results loaded with color-coded analysis!")

        except Exception as e:
            messagebox.showerror("Error", f"Comparison failed: {e}")

    def showComparisonGraph(self):
        if not hasattr(self, 'comparisonData'):
            messagebox.showwarning("Warning", "Please compare CSV files first!")
            return
            
        self.xrdPage.pack_forget()
        self.graphPage.pack(fill=tk.BOTH, expand=True)

        for widget in self.graphCanvasFrame.winfo_children():
            widget.destroy()

        self.fig = Figure(figsize=(10, 5), dpi=100)
        self.comparisonAx1 = self.fig.add_subplot(111)

        self.comparisonCanvas = FigureCanvasTkAgg(self.fig, master=self.graphCanvasFrame)
        self.comparisonCanvas.draw()
        
        self.comparisonToolbar = NavigationToolbar2Tk(self.comparisonCanvas, self.graphCanvasFrame)
        self.comparisonToolbar.update()
        
        self.comparisonCanvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

        self.comparisonAnimIndex = 0
        self.updateComparisonAnimation()

    def updateComparisonAnimation(self):
        tt1, int1, name1 = self.comparisonData["file1"]
        tt2, int2, name2 = self.comparisonData["file2"]
        
        tt1, int1 = np.array(tt1), np.array(int1)
        tt2, int2 = np.array(tt2), np.array(int2)

        totalSteps1 = len(tt1)
        totalSteps2 = len(tt2)
        maxSteps = max(totalSteps1, totalSteps2)

        if self.comparisonAnimIndex < maxSteps:
            stepChunk = max(1, maxSteps // 150)
            currentEnd = min(self.comparisonAnimIndex + stepChunk, maxSteps)

            end1 = min(currentEnd, totalSteps1)
            end2 = min(currentEnd, totalSteps2)

            self.comparisonAx1.clear()

            self.comparisonAx1.plot(tt1[:end1], int1[:end1], label=name1, color="blue", linewidth=1.8)
            self.comparisonAx1.plot(tt2[:end2], int2[:end2], label=name2, color="orange", linewidth=1.8)

            self.comparisonAx1.set_title("XRD Pattern Overlay Comparison (Live)")
            self.comparisonAx1.set_xlabel("2θ (Degrees)")
            self.comparisonAx1.set_ylabel("Relative Intensity (%)")
            self.comparisonAx1.set_ylim(0, 110)
            self.comparisonAx1.grid(True, linestyle="--", alpha=0.6)
            self.comparisonAx1.legend()

            self.comparisonCanvas.draw()

            self.comparisonAnimIndex = currentEnd
            self.hub.mainPage.after(5, self.updateComparisonAnimation)

    def saveGraphImage(self):
        if not hasattr(self, 'fig') or not hasattr(self, 'safeMaterialName') or not hasattr(self, 'resultFolder'):
            messagebox.showerror("Error", "No active graph found to save! Please run simulation first.")
            return
            
        try:
            self.graphFileName = os.path.join(self.resultFolder, f"Graph of {self.safeMaterialName}.png")
            self.fig.savefig(self.graphFileName, dpi=300, bbox_inches='tight')
            messagebox.showinfo("Saved", f"Graph successfully saved to:\n{self.graphFileName}")
        except Exception as e:
            messagebox.showerror("Save Error", f"Could not save graph image: {e}")

    def startLiveAnimation(self):
        self.xrdPage.pack_forget()
        self.graphPage.pack(fill=tk.BOTH, expand=True)

        for widget in self.graphCanvasFrame.winfo_children():
            widget.destroy()

        self.fig = Figure(figsize=(10, 5), dpi=100)
        self.ax1 = self.fig.add_subplot(111)

        self.canvas = FigureCanvasTkAgg(self.fig, master=self.graphCanvasFrame)
        self.canvas.draw()
        
        self.toolbar = NavigationToolbar2Tk(self.canvas, self.graphCanvasFrame)
        self.toolbar.update()
        
        self.canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

        self.updateAnimation()

    def updateAnimation(self):
        totalSteps = len(self.listofTwoTheta)

        if self.animIndex < totalSteps:
            stepChunk = max(1, totalSteps // 150)
            currentEnd = min(self.animIndex + stepChunk, totalSteps)

            twoThetaSlice = self.listofTwoTheta[:currentEnd]
            intensitySlice = self.listofIntensity[:currentEnd]

            self.ax1.clear()
            self.ax1.plot(twoThetaSlice, intensitySlice, color="darkblue", linewidth=1.8)
            self.ax1.set_title(f"XRD Pattern - {self.safeMaterialName} ({self.crystalSystem} / {self.crystalStructure})")
            self.ax1.set_xlabel("2θ (Degrees)")
            self.ax1.set_ylabel("Relative Intensity (%)")
            self.ax1.set_xlim(self.minTwoTheta, self.maxTwoTheta)
            self.ax1.set_ylim(0, 110)
            self.ax1.grid(True, linestyle="--", alpha=0.6)

            self.canvas.draw()
            self.animIndex = currentEnd

            if self.animIndex < totalSteps:
                self.hub.mainPage.after(5, self.updateAnimation)
            else:
                for two_theta_deg, hkl, d_hkl, rel_int in self.detectedPeaks:
                    self.ax1.plot(two_theta_deg, rel_int, 'ro', markersize=4)
                    hkl_label = f"({hkl[0]}{hkl[1]}{hkl[2]})\n{two_theta_deg:.1f}°"
                    self.ax1.annotate(
                        hkl_label, 
                        xy=(two_theta_deg, rel_int), 
                        xytext=(0, 8), 
                        textcoords="offset points", 
                        ha="center", 
                        fontsize=8, 
                        fontweight="bold", 
                        color="red"
                    )

                self.canvas.draw()
                messagebox.showinfo("Success", "XRD simulation completed successfully!")


if __name__ == "__main__":
    mainPage = tk.Tk()
    app = TulparHub(mainPage)
    mainPage.mainloop()
