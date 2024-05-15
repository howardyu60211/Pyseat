import tkinter.dialog
import tkinter.filedialog
import tkinter.messagebox
import pandas as pd
import tkinter

from ttkbootstrap import ttk
from ttkbootstrap.window import Window

from random import randint

import ctypes


class Sorter(Window):
    def __init__(self):
        super().__init__(themename="darkly")
        self.seatRow = None
        self.seatCol = None
        self.seatStatus = None
        self.seatBtn = None
        self.stu_df = pd.DataFrame()

        self.title("座位產生器 v0.1(Beta)")

        # == Input Label == #
        self.InputFrame = ttk.LabelFrame(self, text="學生資料", bootstyle="warning")
        self.tree = ttk.Treeview(self.InputFrame, column=("座號", "姓名"), show='headings')
        self.tree.heading("#1", text="座號")
        self.tree.heading("#2", text="姓名")
        self.tree.grid(column=0, row=0, columnspan=3, padx=5, pady=5, sticky=tkinter.N + tkinter.S)
        self.tree.bind("<Delete>", self.DeleteStu)

        self.chooseBtn = ttk.Button(self.InputFrame, command=self.importData, text="選擇檔案")
        self.chooseBtn.grid(column=0, row=1, pady=5)

        self.SortBtn = ttk.Button(self.InputFrame, text="排序！", command=self.randDisplay, state=tkinter.DISABLED)
        self.SortBtn.grid(column=1, row=1)

        self.exportBtn = ttk.Button(self.InputFrame, text="匯出.csv檔", command=self.exportSeat, state=tkinter.DISABLED)
        self.exportBtn.grid(column=2, row=1)

        self.InputFrame.grid(column=0, row=0, sticky=tkinter.N + tkinter.S, pady=10, padx=20)
        """
        # == output Label == #
        self.OutputFrame = ttk.LabelFrame(self)

        self.outputLabel = ttk.Label(self.OutputFrame, text="Starting...")
        self.outputLabel.pack(fill="both")

        self.OutputFrame.grid(column=0, row=1)
        """
        # == processing Label == #
        self.processingFrame = ttk.LabelFrame(self, text="座位模擬", bootstyle="warning")

        self.rcFrame = ttk.LabelFrame(self.processingFrame, borderwidth=0)

        vCmd = self.register(lambda s: str.isdigit(s) or s == ""), '%P'

        self.rowText = ttk.Label(self.rcFrame, text="行數：")
        self.rowText.grid(column=0, row=0)

        self.rowNum = tkinter.IntVar(value=6)
        self.rowEnter = ttk.Entry(self.rcFrame, validatecommand=vCmd, validate="all", textvariable=self.rowNum,
                                  bootstyle="info")
        self.rowEnter.grid(column=1, row=0)

        self.colText = ttk.Label(self.rcFrame, text="列數：")
        self.colText.grid(column=0, row=1, pady=20, padx=10)

        self.colNum = tkinter.IntVar(value=8)
        self.colEnter = ttk.Entry(self.rcFrame, validatecommand=vCmd, validate="all", textvariable=self.colNum,
                                  bootstyle="info")
        self.colEnter.grid(column=1, row=1)

        self.rcBtn = ttk.Button(self.rcFrame, text="生成表格", command=self.generateSeat)
        self.rcBtn.grid(column=2, row=0, padx=15)

        self.rcBtn = ttk.Button(self.rcFrame, text="清空表格", command=self.clearSeat)
        self.rcBtn.grid(column=2, row=1, padx=15)

        self.rcFrame.pack()

        self.guide = ttk.Label(self.processingFrame, text="講台", bootstyle="info")
        self.guide.pack(fill=tkinter.X, padx=20)

        self.seatFrame = ttk.LabelFrame(self.processingFrame, borderwidth=0)

        self.targetCol = -1
        self.targetRow = -1

        self.protection = False

        self.generateSeat()

        self.processingFrame.grid(padx=20, pady=10, column=1, row=0)

    def exportSeat(self):
        data = []
        for i in range(self.seatRow):
            data.append([])
            for j in range(self.seatCol):
                data[i].append(self.seatBtn[(j, i)]["text"])
        df = pd.DataFrame(data)
        fileType = ".csv"
        exportFilePath = tkinter.filedialog.asksaveasfilename(defaultextension=".csv",
                                                              filetypes=[('逗點分隔檔案', '.csv'),
                                                                         ('Excel 活頁簿', '.xlsx')])
        if not exportFilePath: return
        if exportFilePath.rfind(".") != -1:
            fileType = exportFilePath[exportFilePath.rfind(".") + 1:]

        if fileType == "csv":
            df.to_csv(exportFilePath, header=False, index=False, encoding='UTF-8')
        elif fileType == "xlsx" or fileType == 'xlsm':
            df.to_excel(exportFilePath, header=False, index=False)
        else:
            tkinter.messagebox.showerror("error!", "不支援此種檔案模式！(僅支援excel或csv檔)")

    def changeColor(self, r, c):
        btn = self.seatBtn[(r, c)]
        if self.seatStatus[r][c] == 0:
            btn.config(bootstyle="danger", text=" X ")
            self.seatStatus[r][c] = 1
        elif self.seatStatus[r][c] == 1:
            btn.config(bootstyle="success", text="     ")
            self.seatStatus[r][c] = 0

    def deleteSeat(self, _, c, r):
        if self.seatStatus[c][r] == 2:
            self.seatBtn[(c, r)].config(bootstyle="danger", text=" X ")
            self.seatStatus[c][r] = 1

    def DeleteStu(self, _):
        self.tree.delete(self.tree.focus())

    def importData(self) -> None:
        stu_path = tkinter.filedialog.askopenfilename(filetypes=[('逗點分隔檔案', '.csv'), ('Excel 活頁簿', '.xlsx')])
        if not stu_path:
            return
        fileType = ""
        if stu_path.rfind(".") != -1:
            fileType = stu_path[stu_path.rfind(".") + 1:]
        if fileType == "csv":
            self.stu_df = pd.read_csv(stu_path, header=None)
        elif fileType == "xlsx" or fileType == 'xlsm':
            self.stu_df = pd.read_excel(stu_path, header=None)
        else:
            tkinter.messagebox.showerror("error!", "不支援此種檔案模式！(僅支援excel或csv檔)")
            return

        print(self.stu_df)
        self.displayTreeData()
        self.SortBtn['state'] = tkinter.NORMAL

    def displayChosen(self, _, r, c):
        self.guide["text"] = "已選{} (按delete鍵刪除)".format(self.seatBtn[(c, r)]["text"])

    def randDisplay(self):
        self.stu_df = self.stu_df.sample(frac=1)
        greenNum = 0
        for row in self.seatStatus:
            for e in row:
                if e == 0:
                    greenNum += 1

        stuLen = len(self.tree.get_children())

        if greenNum < stuLen:
            tkinter.messagebox.showerror("Error!",
                                         f"學生數量({stuLen}人)大於座位數量({greenNum}席)！\n請增加座位數或刪除學生。")
            return
        elif greenNum > stuLen:
            conti = tkinter.messagebox.askokcancel("Warning!",
                                                   f"座位數量({greenNum}席)大於學生數量({stuLen}人)！\n將會隨機選擇空位，是否繼續排序？")
            if not conti:
                return

        while stuLen != 0:
            randR = randint(0, self.seatRow - 1)
            randC = randint(0, self.seatCol - 1)
            while self.seatStatus[randC][randR] != 0:
                randR = randint(0, self.seatRow - 1)
                randC = randint(0, self.seatCol - 1)
            self.seatBtn[(randC, randR)].configure(bootstyle="info", text=str(
                self.tree.item(self.tree.get_children()[0])["values"][0]) + " " + str(
                self.tree.item(self.tree.get_children()[0])["values"][1]))
            self.seatBtn[(randC, randR)].bind("<Delete>",
                                              lambda event, rc=randC, rr=randR: self.deleteSeat(event, rc, rr))
            self.seatBtn[(randC, randR)].bind("<ButtonRelease-1>",
                                              lambda event, c=randC, r=randR: self.displayChosen(event, c, r))
            self.seatStatus[randC][randR] = 2
            self.tree.delete(self.tree.get_children()[0])
            stuLen -= 1

        for i in range(self.seatCol):
            for j in range(self.seatRow):
                if self.seatStatus[i][j] == 0:
                    self.seatBtn[(i, j)].config(bootstyle="danger", text=" X ")
                    self.seatStatus[i][j] = 1

        for i in range(self.seatCol):
            for j in range(self.seatRow):
                self.seatBtn[(i, j)].bind("<Button-3>", lambda event, rc=i, rr=j: self.changeSeatClick(event, rc, rr))

        self.SortBtn["state"] = tkinter.DISABLED
        self.exportBtn["state"] = tkinter.NORMAL

    def displayTreeData(self):
        r_set = self.stu_df.to_numpy().tolist()
        for dt in r_set:
            v = [r for r in dt]  # collect the row data as list
            self.tree.insert("", 'end', iid=v[0], values=v)

    def generateSeat(self):
        if self.rowNum.get() <= 0 or self.colNum.get() <= 0: tkinter.messagebox.showerror("行列數須為正整數！")
        if self.rowNum.get() > 30 or self.colNum.get() > 40:
            if (not tkinter.messagebox.askokcancel("Warning!",
                                                   "數量超過建議限制，是否繼續執行？\n(是：繼續執行、取消：撤回動作)")): return

        self.seatFrame.destroy()
        self.seatBtn = {}
        self.seatStatus = []
        self.seatFrame = ttk.LabelFrame(self.processingFrame, borderwidth=0)
        self.seatCol = self.colNum.get()
        self.seatRow = self.rowNum.get()

        for i in range(self.colNum.get()):
            self.seatStatus.append([])
            for j in range(self.rowNum.get()):
                self.seatStatus[i].append(0)
                btn = ttk.Button(self.seatFrame, command=lambda i=i, j=j: self.changeColor(i, j), bootstyle="success",
                                 text="     ")
                self.seatBtn[(i, j)] = btn
                self.seatBtn[(i, j)].grid(column=i, row=j, padx=3, pady=3, sticky=tkinter.W + tkinter.E)

        self.seatFrame.pack(pady=5)

    def clearSeat(self):
        for i in range(self.seatCol):
            for j in range(self.seatRow):
                self.seatBtn[(i, j)].config(bootstyle="success", text="     ")
                self.seatStatus[i][j] = 0

    def changeSeatClick(self, _, tc: int, tr: int):
        if self.targetRow == -1 or self.targetCol == -1:
            self.targetCol = tc
            self.targetRow = tr
            if self.seatStatus[tc][tr] == 0:
                self.seatBtn[(tc, tr)].config(bootstyle="success-outline")
            elif self.seatStatus[tc][tr] == 1:
                self.seatBtn[(tc, tr)].config(bootstyle="danger-outline")
            elif self.seatStatus[tc][tr] == 2:
                self.seatBtn[(tc, tr)].config(bootstyle="info-outline")
            self.guide["text"] = "已選擇 " + self.seatBtn[(tc, tr)]["text"] + "，與誰交換位置？"
        else:
            tempt = str(self.seatBtn[(self.targetCol, self.targetRow)].cget("text"))
            temps = self.seatStatus[self.targetCol][self.targetRow]

            self.seatBtn[(self.targetCol, self.targetRow)].config(text=str(self.seatBtn[(tc, tr)].cget("text")))
            if self.seatStatus[tc][tr] == 0:
                self.seatBtn[(self.targetCol, self.targetRow)].config(bootstyle="success")
            elif self.seatStatus[tc][tr] == 1:
                self.seatBtn[(self.targetCol, self.targetRow)].config(bootstyle="danger")
            elif self.seatStatus[tc][tr] == 2:
                self.seatBtn[(self.targetCol, self.targetRow)].config(bootstyle="info")
            self.seatStatus[self.targetCol][self.targetRow] = self.seatStatus[tc][tr]

            self.seatBtn[(tc, tr)].config(text=tempt)
            if temps == 0:
                self.seatBtn[(tc, tr)].config(bootstyle="success")
            elif temps == 1:
                self.seatBtn[(tc, tr)].config(bootstyle="danger")
            elif temps == 2:
                self.seatBtn[(tc, tr)].config(bootstyle="info")
            self.seatStatus[tc][tr] = temps

            if self.seatStatus[self.targetCol][self.targetRow] == 0:
                self.seatBtn[(self.targetCol, self.targetRow)].config(bootstyle="success")
            elif self.seatStatus[self.targetCol][self.targetRow] == 1:
                self.seatBtn[(self.targetCol, self.targetRow)].config(bootstyle="danger")
            elif self.seatStatus[self.targetCol][self.targetRow] == 2:
                self.seatBtn[(self.targetCol, self.targetRow)].config(bootstyle="info")
            self.targetCol = -1
            self.targetRow = -1

            self.guide["text"] = "講台"


if __name__ == "__main__":
    ctypes.windll.shcore.SetProcessDpiAwareness(2)
    sorter = Sorter()
    sorter.mainloop()
