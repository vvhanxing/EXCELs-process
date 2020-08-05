#!/usr/bin/env python

import wx
from DataViewModel5 import *
import find_ID_info_pandas_9
files_name_text_1= []
files_name_text_2= []
files_name_text_3= []
#----------------------------------------------------------------------

ID_CopyBtn      = wx.NewIdRef()
ID_PasteBtn     = wx.NewIdRef()
ID_BitmapBtn    = wx.NewIdRef()
ID_CreatDataBtn = wx.NewIdRef()
ID_NextBtn    = wx.NewIdRef()
ID_ClearBtn    = wx.NewIdRef()

#----------------------------------------------------------------------

class ClipTextPanel(wx.Panel):
    def __init__(self, parent, log):
        wx.Panel.__init__(self, parent, -1)
        self.log = log

        sizer = wx.BoxSizer(wx.VERTICAL)
        sizer.Add(
            wx.StaticText(
                self, -1, "Copy/Paste text to/from\n"
                "this window and other apps"
                ),
            0, wx.EXPAND|wx.ALL, 2
            )

        self.text = wx.TextCtrl(self, -1, "", style=wx.TE_MULTILINE|wx.HSCROLL)
        sizer.Add(self.text, 1, wx.EXPAND)

        hsz = wx.BoxSizer(wx.HORIZONTAL)
        hsz.Add(wx.Button(self, ID_CopyBtn, " Copy "), 1, wx.EXPAND|wx.ALL, 2)
        hsz.Add(wx.Button(self, ID_PasteBtn, " Paste "), 1, wx.EXPAND|wx.ALL, 2)
        sizer.Add(hsz, 0, wx.EXPAND)
        sizer.Add(
            wx.Button(self, ID_BitmapBtn, " Copy Bitmap "),
            0, wx.EXPAND|wx.ALL, 2
            )

        self.Bind(wx.EVT_BUTTON, self.OnCopy, id=ID_CopyBtn)
        self.Bind(wx.EVT_BUTTON, self.OnPaste, id=ID_PasteBtn)
        self.Bind(wx.EVT_BUTTON, self.OnCopyBitmap, id=ID_BitmapBtn)

        self.SetSizer(sizer)


    def OnCopy(self, evt):
        self.do = wx.TextDataObject()
        self.do.SetText(self.text.GetValue())
        if wx.TheClipboard.Open():
            wx.TheClipboard.SetData(self.do)
            wx.TheClipboard.Close()
        else:
            wx.MessageBox("Unable to open the clipboard", "Error")


    def OnPaste(self, evt):

        success = False
        do = wx.TextDataObject()
        if wx.TheClipboard.Open():
            success = wx.TheClipboard.GetData(do)
            wx.TheClipboard.Close()

        if success:
            self.text.SetValue(do.GetText())
            #input_file_names = os.listdir("data")
            #Panel_data(input_file_names)
        else:
            wx.MessageBox(
                "There is no data in the clipboard in the required format",
                "Error"
                )


    def OnCopyBitmap(self, evt):
        dlg = wx.FileDialog(self, "Choose a bitmap to copy", wildcard="*.bmp")

        if dlg.ShowModal() == wx.ID_OK:
            bmp = wx.Bitmap(dlg.GetPath(), wx.BITMAP_TYPE_BMP)
            bmpdo = wx.BitmapDataObject(bmp)
            if wx.TheClipboard.Open():
                wx.TheClipboard.SetData(bmpdo)
                wx.TheClipboard.Close()

                wx.MessageBox(
                    "The bitmap is now in the Clipboard.  Switch to a graphics\n"
                    "editor and try pasting it in..."
                    )
            else:
                wx.MessageBox(
                    "There is no data in the clipboard in the required format",
                    "Error"
                    )

        dlg.Destroy()

#----------------------------------------------------------------------

class OtherDropTarget(wx.DropTarget):
    def __init__(self, window, log):
        wx.DropTarget.__init__(self)
        self.log = log
        self.do = wx.FileDataObject()
        self.SetDataObject(self.do)

    def OnEnter(self, x, y, d):
        self.log.WriteText("OnEnter: %d, %d, %d\n" % (x, y, d))
        return wx.DragCopy

    #def OnDragOver(self, x, y, d):
    #    self.log.WriteText("OnDragOver: %d, %d, %d\n" % (x, y, d))
    #    return wx.DragCopy

    def OnLeave(self):
        self.log.WriteText("OnLeave\n")

    def OnDrop(self, x, y):
        self.log.WriteText("OnDrop: %d %d\n" % (x, y))
        return True

    def OnData(self, x, y, d):
        self.log.WriteText("OnData: %d, %d, %d\n" % (x, y, d))
        self.GetData()
        self.log.SetLabel("%s\n" % self.do.GetFilenames())
        return d


class MyFileDropTarget_1(wx.FileDropTarget):
    def __init__(self, window, log):
        wx.FileDropTarget.__init__(self)
        self.window = window
        self.log = log


    def OnDropFiles(self, x, y, filenames):
        global files_name_text_1
        txt = "\n%d file(s) dropped at %d,%d:\n" % (len(filenames), x, y)
        if len(filenames)==1:  
            files_name_text_1 .extend( ['',filenames[0],''])
        else :
            files_name_text_1.extend(filenames)
        txt += '\n'.join(files_name_text_1)
        self.window.SetLabel(txt)
        return True


    def OnClearFiles(self, x, y, filenames):
        global files_name_text_1
        txt = "\n%d file(s) dropped at %d,%d:\n" % (len(filenames), x, y)
        files_name_text_1 = []
        self.window.SetLabel(txt)
        return True



class MyFileDropTarget_2(wx.FileDropTarget):
    def __init__(self, window, log):
        wx.FileDropTarget.__init__(self)
        self.window = window
        self.log = log


    def OnDropFiles(self, x, y, filenames):
        global files_name_text_2
        txt = "\n%d file(s) dropped at %d,%d:\n" % (len(filenames), x, y)
        if len(filenames)==1:  
            files_name_text_2 .extend( ['',filenames[0],''])
        else :
            files_name_text_2.extend(filenames)
        txt += '\n'.join(files_name_text_2)
        self.window.SetLabel(txt)
        return True


    def OnClearFiles(self, x, y, filenames):
        global files_name_text_2
        txt = "\n%d file(s) dropped at %d,%d:\n" % (len(filenames), x, y)
        files_name_text_2 = []
        self.window.SetLabel(txt)
        return True


class MyFileDropTarget_3(wx.FileDropTarget):
    def __init__(self, window, log):
        wx.FileDropTarget.__init__(self)
        self.window = window
        self.log = log


    def OnDropFiles(self, x, y, filenames):
        global files_name_text_3
        txt = "\n%d file(s) dropped at %d,%d:\n" % (len(filenames), x, y)
        if len(filenames)==1:  
            files_name_text_3 .extend( ['',filenames[0],''])
        else :
            files_name_text_3.extend(filenames)
        txt += '\n'.join(files_name_text_3)
        self.window.SetLabel(txt)
        return True


    def OnClearFiles(self, x, y, filenames):
        global files_name_text_3
        txt = "\n%d file(s) dropped at %d,%d:\n" % (len(filenames), x, y)
        files_name_text_3 = []
        self.window.SetLabel(txt)
        return True



class MyTextDropTarget(wx.TextDropTarget):
    def __init__(self, window, log):
        wx.TextDropTarget.__init__(self)
        self.window = window
        self.log = log

    def OnDropText(self, x, y, text):
        self.window.SetLabel("(%d, %d)\n%s\n" % (x, y, text))
        return True

    def OnDragOver(self, x, y, d):
        return wx.DragCopy





class FileDropPanel_1(wx.Panel):
    def __init__(self, parent, log):
        wx.Panel.__init__(self, parent, -1)

        sizer = wx.BoxSizer(wx.VERTICAL)
    
        sizer.Add(
            wx.StaticText(self, -1, " \nDrag FIND files here:"),
            0, wx.EXPAND|wx.ALL, 2
            )

        self.text = wx.StaticText(self, -1, "", style=wx.ST_NO_AUTORESIZE|wx.BORDER_SIMPLE)
        self.text.SetWindowVariant(wx.WINDOW_VARIANT_SMALL)

        dt = MyFileDropTarget_1(self, log)
        self.text.SetDropTarget(dt)
        sizer.Add(self.text, 1, wx.EXPAND|wx.ALL, 5)

        self.SetAutoLayout(True)
        self.SetSizer(sizer)


    def SetLabel(self, text):
        self.text.SetLabel(text)


class FileDropPanel_2(wx.Panel):
    def __init__(self, parent, log):
        wx.Panel.__init__(self, parent, -1)

        sizer = wx.BoxSizer(wx.VERTICAL)
    
        sizer.Add(
            wx.StaticText(self, -1, " \nDrag DATA files here:"),
            0, wx.EXPAND|wx.ALL, 2
            )

        self.text = wx.StaticText(self, -1, "", style=wx.ST_NO_AUTORESIZE|wx.BORDER_SIMPLE)
        self.text.SetWindowVariant(wx.WINDOW_VARIANT_SMALL)

        dt = MyFileDropTarget_2(self, log)
        self.text.SetDropTarget(dt)
        sizer.Add(self.text, 1, wx.EXPAND|wx.ALL, 5)

        self.SetAutoLayout(True)
        self.SetSizer(sizer)


    def SetLabel(self, text):
        self.text.SetLabel(text)


class FileDropPanel_3(wx.Panel):
    def __init__(self, parent, log):
        wx.Panel.__init__(self, parent, -1)

        sizer = wx.BoxSizer(wx.VERTICAL)
    
        sizer.Add(
            wx.StaticText(self, -1, " \nDrag MODEL files here:"),
            0, wx.EXPAND|wx.ALL, 2
            )

        self.text = wx.StaticText(self, -1, "", style=wx.ST_NO_AUTORESIZE|wx.BORDER_SIMPLE)
        self.text.SetWindowVariant(wx.WINDOW_VARIANT_SMALL)

        dt = MyFileDropTarget_3(self, log)
        self.text.SetDropTarget(dt)
        sizer.Add(self.text, 1, wx.EXPAND|wx.ALL, 5)

        self.SetAutoLayout(True)
        self.SetSizer(sizer)


    def SetLabel(self, text):
        self.text.SetLabel(text)
#----------------------------------------------------------------------
#----------------------------------------------------------------------

class TestPanel(wx.Panel):
    def __init__(self, parent, log):
        wx.Panel.__init__(self, parent, -1)

        self.SetAutoLayout(True)
        outsideSizer = wx.BoxSizer(wx.VERTICAL)

        msg = "Drag files"
        text = wx.StaticText(self, -1, "", style=wx.ALIGN_CENTRE)
        text.SetFont(wx.Font(24, wx.FONTFAMILY_SWISS, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD, False))
        text.SetLabel(msg)      

        w,h = text.GetTextExtent(msg)
        text.SetSize(wx.Size(w,h+1))
        text.SetForegroundColour(wx.BLUE)
        outsideSizer.Add(text, 0, wx.EXPAND|wx.ALL, 5)
        outsideSizer.Add(wx.StaticLine(self, -1), 0, wx.EXPAND)

        inSizer = wx.BoxSizer(wx.HORIZONTAL)
        #input_file_names = os.listdir("data")
        #print(files_name_text_1)
        #global files_name_text_1
        #print(input_file_names)
        #data = Panel_data(files_name_text_1,select_list)
        #inSizer.Add(TestPanel_excel(self, log, data=data), 1, wx.EXPAND)
        inSizer.Add(FileDropPanel_1(self, log), 1, wx.EXPAND)
        inSizer.Add(FileDropPanel_2(self, log), 1, wx.EXPAND)
        inSizer.Add(FileDropPanel_3(self, log), 1, wx.EXPAND)
        #inSizer = wx.BoxSizer(wx.HORIZONTAL)
        #input_file_names = os.listdir("data")
        #print(files_name_text_1)
        #global files_name_text_1
        #print(input_file_names)
        #data = Panel_data(files_name_text_1,select_list)
        #inSizer.Add(TestPanel_excel(self, log, data=data), 1, wx.EXPAND)
        #inSizer.Add(FileDropPanel(self, log), 1, wx.EXPAND)
        #outsideSizer.Add(FileDropPanel(self, log), 1, wx.EXPAND)

        creatData_button = wx.Button(self, ID_CreatDataBtn, "CreatData")
        self.Bind(wx.EVT_BUTTON, self.On_creatData_button, id=ID_CreatDataBtn)

        next_button = wx.Button(self, ID_NextBtn, "Next")
        self.Bind(wx.EVT_BUTTON, self.On_next_button, id=ID_NextBtn)

        #clear_button = wx.Button(self, ID_ClearBtn, "Clear")

        #btnbox = wx.BoxSizer(wx.HORIZONTAL)vertical 
        btnbox = wx.BoxSizer(wx.VERTICAL)
        btnbox.Add(creatData_button, 0, wx.TOP|wx.BOTTOM, 5)
        btnbox.Add(next_button, 0, wx.TOP|wx.BOTTOM, 5)
        inSizer.Add(btnbox, 0, wx.TOP|wx.BOTTOM, 5)

        outsideSizer.Add(inSizer, 2, wx.EXPAND)
        self.SetSizer(outsideSizer)


####################
    def On_creatData_button(self, evt):
        #win = MyFrame(self, -1, "This is a wx.Frame", size=(350, 200),
                  #style = wx.DEFAULT_FRAME_STYLE)
        #win.Show(True)
        global files_name_text_1
        global files_name_text_2
        global files_name_text_3
        #print(files_name_text_1)
        files_name_text_1 = ["\\".join(i.split("\\")) for i in files_name_text_1 if i!=""]
        files_name_text_2 = ["\\".join(i.split("\\")) for i in files_name_text_2 if i!=""]
        files_name_text_3 = ["\\".join(i.split("\\")) for i in files_name_text_3 if i!=""]
        #find_ID_info_pandas_6.creat_all_data_xlsx(files_name_text_1)
        print("files_name_text_1",files_name_text_1)
        print("files_name_text_2",files_name_text_2)
        print("files_name_text_3",files_name_text_3)
        find_ID_info_pandas_9.creat_all_data_xlsx(files_name_text_2)
        print("------------finish-----------------")
        
    def On_next_button(self, evt):
        #win = MyFrame(self, -1, "This is a wx.Frame", size=(350, 200),
                  #style = wx.DEFAULT_FRAME_STYLE)
        #win.Show(True)
        global files_name_text_1
        global files_name_text_2
        global files_name_text_3
        #print(files_name_text_1)
        files_name_text_1 = ["\\".join(i.split("\\")) for i in files_name_text_1 if i!=""]
        files_name_text_2 = ["\\".join(i.split("\\")) for i in files_name_text_2 if i!=""]
        files_name_text_3 = ["\\".join(i.split("\\")) for i in files_name_text_3 if i!=""]
        #find_ID_info_pandas_6.creat_all_data_xlsx(files_name_text_1)
        print("files_name_text_1",files_name_text_1)
        print("files_name_text_2",files_name_text_2)
        print("files_name_text_3",files_name_text_3)
        #find_ID_info_pandas_9.creat_all_data_xlsx(files_name_text_2)

        #############
        
        #input_excel_name = files_name_text_1[0]
        #model_excel_name = files_name_text_3[-1]
        #save_name = 'result.xlsx'
        #find_ID_info_pandas_9.save_output_excel(input_excel_name,model_excel_name,save_name)
        print("finish")
        ######################
        #input_file_names = os.listdir("data")
        #global select_list
        #select_list = process_excel_data.get_select_txt()

        
        Ordered_category_names = find_ID_info_pandas_9.get_category_names( files_name_text_3[0])#建立有序空字典，顺序是模板的标题名
        select_list = [   find_ID_info_pandas_9.get_standard(n) for n in Ordered_category_names] #把model中标题标准化放入新表中
        

        data =  Panel_data(files_name_text_2,select_list)
        frm = wx.Frame(None, title="DataViewModel sample", size=(700,500))
        pnl = TestPanel_excel(frm, sys.stdout,files_name_text=files_name_text_2 ,input_file_name=files_name_text_1[-1],model_excel_name=files_name_text_3[-1],select_list=select_list, data=data)
        frm.Show()
#####################        





#----------------------------------------------------------------------

def runTest(frame, nb, log):
    win = TestPanel(nb, log)
    return win

#----------------------------------------------------------------------


overview = """\
<html>
<body>
This demo shows some examples of data transfer through clipboard or
drag and drop. In wxWindows, these two ways to transfer data (either
between different applications or inside one and the same) are very
similar which allows to implement both of them using almost the same
code - or, in other words, if you implement drag and drop support for
your application, you get clipboard support for free and vice versa.
<p>
At the heart of both clipboard and drag and drop operations lies the
wxDataObject class. The objects of this class (or, to be precise,
classes derived from it) represent the data which is being carried by
the mouse during drag and drop operation or copied to or pasted from
the clipboard. wxDataObject is a "smart" piece of data because it
knows which formats it supports (see GetFormatCount and GetAllFormats)
and knows how to render itself in any of them (see GetDataHere). It
can also receive its value from the outside in a format it supports if
it implements the SetData method. Please see the documentation of this
class for more details.
<p>
Both clipboard and drag and drop operations have two sides: the source
and target, the data provider and the data receiver. These which may
be in the same application and even the same window when, for example,
you drag some text from one position to another in a word
processor. Let us describe what each of them should do.
</body>
</html>
"""


if __name__ == '__main__':
    import sys,os
    import run
    run.main(['', os.path.basename(sys.argv[0])] + sys.argv[1:])

