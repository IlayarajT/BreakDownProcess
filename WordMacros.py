class WordMacros:
    def AcceptTrackChange(self, doc):
        doc.ActiveDocument.AcceptAllRevisions()

    def CheckDocLocked(self, doc):
        protected = False
        isProtected = doc.ProtectionType
        if isProtected == 3:
            protected = True
        return protected

    def setWebView(self, doc):
        doc.ActiveWindow.View.Type = wdWebView

    def LoadSageStyles(self, doc):
        doc.ActiveDocument.CopyStylesFromTemplate("SupportingFiles/SAGE_styles.dotx")
