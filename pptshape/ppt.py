import pywintypes
import win32com.client, win32com.client.gencache, win32com.gen_py

def open(filename):
    try:
        return PPTShape(filename)
    except pywintypes.com_error as exc:
        if exc.hresult == -2147221005: # Invalid prog-id
            return None
        raise

class PPTShape:
    def __init__(self, filename):
        self.ppt = win32com.client.gencache.EnsureDispatch("PowerPoint.Application") 
        #self.ppt.Visible = 0
        self.filename = filename
        self.presentation = self.ppt.Presentations.Open(self.filename)

    def quit(self):
        self.ppt.Quit()
        self.ppt = None

    def shapes(self):
        for slide in self.presentation.Slides:
            for shape in slide.Shapes:
                yield shape

    def findShape(self, name):
        for shape in self.shapes():
            if shape.Title == name:
                return shape
    
    def saveShape(self, name, filename):
        shape = self.findShape(name)
        if not shape:
            raise ValueError(
                    "Shape '{}' doesn't found in {}".format(name, self.filename))
        #ppRelativeToSlide
        #ppClipRelativeToSlide
        #ppScaleToFit
        #ppScaleXY
        if shape:
            SCALE = 4   # Expand size by 4 to texts to be anti-aliased.

            # ScaleWidth and ScaleHeight are dimentions of slide in ppScaleXY
            # mode.
            w = self.presentation.SlideMaster.Width 
            h = self.presentation.SlideMaster.Height

            shape.Export(filename, 
                Filter=win32com.client.constants.ppShapeFormatPNG,
                ScaleWidth=w*SCALE, ScaleHeight=h*SCALE,
                ExportMode= win32com.client.constants.ppScaleXY)


