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

    def findShapeByIndex(self, name):
        # 'name' is a special shape name as index number(s):
        #  1. '#n' ... whole running number
        #     speify index number of shape in the document.
        #  2. '#m.n' ... running number in the slide
        #     specify slide number 'm' and shape number 'n'
        #     'm' is index number of slide (page) in the document.
        #     'n' is index number of shape in specified slide 'm'.
        assert name.startswith('#')
        nums = name[1:].split('.')
        if len(nums) not in [1, 2]:
            raise ValueError('invalid format of shape index')
        nums = map(int, nums)
        try:
            if len(nums) == 2:
                # '#m.n' format
                slide = self.presentation.Slides.Item(nums[0])
                shape = slide.Shapes.Item(nums[1])
            else:
                # '#n' format
                shapes = list(self.shapes())
                shape = shapes[nums[0] - 1]
        except Exception, ex:
            raise ValueError('nonexistent shape number: ' + name + '\n' + str(ex))
        try:
            t = shape.Title # can be error on PP2007
        except AttributeError:
            t = ''
        #print 'extracting shape %s: %s' % (name, t)
        return shape
            

    def findShape(self, name):
        if name.startswith('#'):
            return self.findShapeByIndex(name)
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


