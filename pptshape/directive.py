import os
from docutils.parsers.rst import directives
from docutils.parsers.rst.directives import images
try:
    from . import ppt
except ImportError:
    ppt = None

class PPTShape(images.Image):
    option_spec = images.Image.option_spec.copy()
    option_spec['pptfilename'] = directives.unchanged_required
    option_spec['shapename'] = directives.unchanged_required

    def run(self):
        if ppt:
            dirname = os.path.split(self.state.document.current_source)[0]
            filename = os.path.join(dirname, self.options['pptfilename'])
            shapename = self.options['shapename']
            imagename = os.path.join(dirname, self.arguments[0])
            pptfile = ppt.open(filename)
            if pptfile:
                pptfile.saveShape(shapename, imagename)
            pptfile.quit()

        return super().run()


def setup(app):
    app.add_directive('ppt-shape', PPTShape)
