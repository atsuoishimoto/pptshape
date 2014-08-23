import itertools
import re
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
    VER_PP2007 = (12, 0)

    def __init__(self, filename):
        self.ppt = win32com.client.gencache.EnsureDispatch("PowerPoint.Application") 
        self.filename = filename

        # check for PPT version
        pptver = tuple(int(n) for n in self.ppt.Version.split('.'))
        if pptver <= self.VER_PP2007:
            self.ppt.Visible = 1 # need to open with PP2007
        self.presentation = self.ppt.Presentations.Open(self.filename)

    def quit(self):
        self.ppt.Quit()
        self.ppt = None

    def shapes(self):
        for slide in self.presentation.Slides:
            for shape in slide.Shapes:
                yield shape

    def extractByIndex(self, idx):
        return next(itertools.islice(self.shapes(), idx+1))

    def extractByPageAndIndex(self, page, idx):
        slide = self.presentation.Slides.Item(page)
        return slide.Shapes.Item(idx)

    def extractByName(self, name):
        for shape in self.shapes():
            if shape.Title == name:
                return shape
    
    def extractShape(self, name):
        
        if name.startswith('#'):
            m = re.match(r'^#([0-9]+)(\.[0-9]+)?$', name.strip())
            if not m:
                raise ValueError('Invalid shape index: %s', name)

            if m.lastindex == 1:
                idx = int(m.group(1))
                if idx <= 0:
                    raise ValueError('Invalid shape index starts from 1: %s', name)
                return self.extractByIndex(idx)
            else:
                page = int(m.group(1))
                idx = int(m.group(2)[1:])
                if page <= 0 or idx <= 0:
                    raise ValueError('Invalid shape index starts from 1: %s', name)
                return self.extractByPageAndIndex(page, idx)

        return self.extractByName(name)

    def saveShape(self, name, filename):
        shape = self.extractShape(name)
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


