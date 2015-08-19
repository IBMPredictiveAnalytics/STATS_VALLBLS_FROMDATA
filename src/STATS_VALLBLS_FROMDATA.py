#/***********************************************************************
# * Licensed Materials - Property of IBM 
# *
# * IBM SPSS Products: Statistics Common
# *
# * (C) Copyright IBM Corp. 2014
# *
# * US Government Users Restricted Rights - Use, duplication or disclosure
# * restricted by GSA ADP Schedule Contract with IBM Corp. 
# ************************************************************************/

from __future__ import with_statement
import random, os, tempfile, textwrap, codecs, re, locale
import spss, spssaux, spssdata, textwrap

"""STATS VALLBLS FROMDATA extension command"""

__author__ =  'IBM SPSS, JKP'
__version__=  '1.0.1'

# history
# 20-jan-2013 Original version

helptext = """STATS VALLLBLS FROMDATA
[/HELP]
* indicates default value.

This command creates value labels for a set of variables
using values of other variables for the labels.  If x is a
variable having values 1,2,3 and xlabel is a variable having
values 'a', 'b', 'c', values labels for x are created as
1 'a'
2 'b'
3 'c'

Syntax:
STATS VALLBLS FROMDATA 
    VARIABLES = variable list or VARPATTERN = "regular expression"
    LBLVARS = string variable list or LBLPATTERN = "regular expression"
OPTIONS
    VARSPERPASS = integer

OUTPUT SYNTAX = "filespec"

/HELP displays this help and does nothing else.

Example:
STATS VALLBLS FROMDATA VARIABLES = x1 TO x5
    LBLVARS = lbl1 to lbl5.
    
VARIABLES lists the variables for which value labels
should be produced.  TO is supported, but ALL would not make sense here.
VARPATTERN is a regular expression in quotes that maps
to the variables to be labelled.  Specify either VARIABLES
or VARPATTERN but not both.  See below for some pattern examples.

LBLVARS lists the variables containing the labels corresponding
to the variables to be labelled.  TO is supported but ALL would not make sense.
Numeric variables are automatically excluded.
LBLPATTERN lists a regular expression in quotes that
expands to the list of label variables.
The number of label variables can be one to apply the same set of labels
to all the selected variables or as many as there are input variables.
If a label variable value is blank, no
value label is generated from it.

VARSPERPASS specifies how many variables should be processed
on a single data pass.  The default is 20.  The size of the
intermediate dataset created by this procedure grows exponentially
with the number of variables.

Value labels are checked for conflicts, i.e., two different labels
for the same value of a variable.  MAXCONFLICTS specifies the
maximum number of conflicts to report across all the variables.
The default value is 100.

REPORTDUPS specifies whether or not to report whether
two or more value labels for a variable are identical.

MAXDUPS specifies the number of duplicates reported.  The
default is 100.
"""

import spss, spssaux, spssdata
from extension import Template, Syntax, processcmd

def dolabels(variables=None, varpattern=None,
    lblvars=None, lblpattern=None, execute=True,
    varsperpass=20, syntax=None):
    """Execute STATS VALLBLS FROMDATA"""
    
# debugging
    # makes debug apply only to the current thread
    #try:
        #import wingdbstub
        #if wingdbstub.debugger != None:
            #import time
            #wingdbstub.debugger.StopDebug()
            #time.sleep(1)
            #wingdbstub.debugger.StartDebug()
        #import thread
        #wingdbstub.debugger.SetDebugThreads({thread.get_ident(): 1}, default_policy=0)
        ## for V19 use
        ###    ###SpssClient._heartBeat(False)
    #except:
        #pass
    try:
        vardict = spssaux.VariableDict(caseless=True)
    except:
        raise ValueError(_("""This command requires a  newer version the spssaux module.  \n
It can be obtained from the SPSS Community website (www.ibm.com/developerworks/spssdevcentral)"""))
    
    varstolabel = resolve(vardict, _("variables to label"), variables, varpattern, stringonly=False)
    labelvars = resolve(vardict, _("label variables"), lblvars, lblpattern, stringonly=True)
    if len(varstolabel) == 0 or len(labelvars) == 0:
        raise ValueError(_("""No variables to label or no labelling variables were specified.
If a pattern was used, it may not have matched any variables."""))
    if len(labelvars) > 1 and len(labelvars) != len(varstolabel):
        raise ValueError(_("The number of label variables is different from the number of variables to label"))
    if min([vardict[item].VariableType for item in labelvars]) == 0:
        raise ValueError(_("""The label variables must all have type string"""))
    dsname = spss.ActiveDataset()
    if dsname == "*":
        raise ValueError(_("""The active dataset must have a dataset name in order to use this procedure"""))
    if syntax:
        syntax = syntax.replace("\\", "/")
        syntax = FileHandles().resolve(syntax)
        
    mkvl = Mkvls(varstolabel, labelvars, varsperpass, execute, syntax, vardict)
    
    for i in range(0, len(varstolabel), varsperpass):
        spss.Submit("""DATASET ACTIVATE %s""" % dsname)
        mkvl.doaggr(i)
    spss.Submit("""DATASET ACTIVATE %s""" % dsname)    
    labelsyntax = mkvl.dolabels()
    if labelsyntax and execute:
        spss.Submit(labelsyntax)
    mkvl.report(labelsyntax)
    if labelsyntax and syntax:
        writesyntax(labelsyntax, syntax, mkvl)


def writesyntax(labelsyntax, syntax, mkvl):
    if mkvl.unicodemode:
        inputencoding = "unicode_internal"
        outputencoding = "utf_8_sig"
    else:
        inputencoding = locale.getlocale()[1]
        outputencoding = inputencoding
    with codecs.EncodedFile(codecs.open(syntax, "wb"), inputencoding, outputencoding) as f:            
        for line in labelsyntax:
            f.write(line + "\n")

def resolve(vardict, itemtype, varlist, pattern, stringonly):
    """Return validated list of variables
    
    itemtype identifies the input description for error message purposes
    varlist is a sequence of variable names or None
    pattern is a regular expression or None
    vardict is a variable dictionary
    stringonly = True excludes numeric variables from pattern matches
    
    If pattern is used, list is returned in SPSS dictionary order"""
    
    if (varlist is None and pattern is None) or\
       (varlist is not None and pattern is not None):
        raise ValueError(_("Either a variable list or a pattern must be specified but not both: %s") % itemtype)
    
    if varlist:
        return vardict.expand(varlist)
    else:
        if stringonly:
            selectedvars = vardict.variablesf(pattern=pattern, variableType="string")
        else:
            selectedvars = vardict.variablesf(pattern=pattern)
        varsinorder = sorted([(vardict[vname].VariableIndex, vname)\
            for vname in selectedvars])
        return [vname for (index, vname) in varsinorder]    



class Mkvls(object):
    """Make Value Labels"""
    
    aggrtemplate = """DATASET DECLARE %s.
AGGREGATE /OUTFILE=%s /BREAK %s"""

    def __init__(self, varstolabel, labelvars, varsperpass, execute, 
        syntax, vardict):
        
        attributesFromDict(locals())
        self.conflicts = {}   # a dictionary of sets of conflicts keyed  by varname
        self.duplabels = {}   # a dictionary of sets of duplicate labels keyed by varname
        self.vlabels = {}     # a dictionary of sets of value label/value pairs keyed by varname
        self.values = {}      # a dictionary of sets of values keyed by varname
        self.labelusage = {}  # a dictionary of dictionaries indexed by varnme and label text
        
        # results are accumulated across data passes
        for v in varstolabel:
            self.conflicts[v] = set()
            #self.duplabels[v] = 0
            self.duplabels[v] = set()
            self.vlabels[v] = set()
            self.values[v] = set()
            self.labelusage[v] = {}
        self.aggrdsname = mkrandomname(sav=False)
        self.unicodemode = spss.PyInvokeSpss.IsUTF8mode()
        if self.unicodemode:
            self.ec = codecs.getencoder("utf_8")   # must figure string length in bytes of utf-8
        
    def doaggr(self, doindex):
        """create an aggregate dataset and tally values
        
        doindex is the index into varstolabel at which to start"""
        
        vtl = self.varstolabel[doindex:doindex+self.varsperpass]
        vtllen = len(vtl)
        if len(self.labelvars) == 1:
            lbls = self.labelvars
            lastlbl = vtllen + 1
        else:
            lbls = self.labelvars[doindex:doindex+self.varsperpass]
            lastlbl = 3 * vtllen - 1
        brkvarlist = "\n".join(textwrap.wrap(" ".join(vtl), width=100))
        outvars = ["/min_%s=MIN(%s)/max_%s=MAX(%s)" % (mkrandomname(), v, mkrandomname(), v) for v in lbls]
        aggrcmd = Mkvls.aggrtemplate % (self.aggrdsname, self.aggrdsname, brkvarlist) + "\n".join(outvars)
        spss.Submit(aggrcmd)
        spss.Submit("DATASET ACTIVATE %s" % self.aggrdsname)
        
        # for each variable, build label information based on data
        # AGGREGATE dataset structure:
        # var1value, var2value,..., min(text lbl1), max(text lbl1), min(text lbl2), max(text lbl2)...
        # but if only one label set, only one pair of label aggregates is produced
        # user missing values are exposed and subject to labelling
        
        curs = spssdata.Spssdata(names=False, convertUserMissing=False)
        for case in curs:
            for v, vname in enumerate(vtl):
                value = case[v]
                minlbl = self.truncate(case[min(vtllen + v*2, lastlbl-1)], 120).rstrip()
                maxlbl = self.truncate(case[min(vtllen + v*2 + 1, lastlbl)], 120).rstrip()
                # more than one label for the same value?
                if minlbl != maxlbl and (minlbl != "" and minlbl is not None):
                    self.conflicts[vname].add(value)
                # ignore empty or missing labels
                if maxlbl != "" and maxlbl is not None:
                    # if the value has already been seen but with a different label, it's a conflict
                    if value in self.values[vname] and not (value, maxlbl) in self.vlabels[vname]:
                        self.conflicts[vname].add(value)
                    else:
                        self.vlabels[vname].add((value, maxlbl))  # first one wins
                        self.values[vname].add(value)
                        # tally instances where the same label used for different value
                        # need to see whether labels has been assigned to a different value
                        previousvalue =  self.labelusage[vname].get(maxlbl, None)
                        if previousvalue is not None and value != previousvalue:
                            ###self.duplabels[vname] = self.duplabels[vname] + 1
                            self.duplabels[vname].add(maxlbl)
                        self.labelusage[vname][maxlbl] = value

        curs.CClose()
        spss.Submit("DATASET CLOSE %s" % self.aggrdsname)
                    
    def dolabels(self):
        """generate, save, and run labelling syntax and write reports"""
        
        vlsyntax = []
        for k,v in sorted(self.vlabels.items()):
            vlsyntax.append(self.makevls(k,v))
        return vlsyntax
            
    def makevls(self, varname, vlinfo):
        """Return value label syntax
        
        varname is the variable to which the syntax applies
        vlinfo is the set of duples of (value, label)"""
        
        isstring = self.vardict[varname].VariableType > 0
        vls = []
        for value, label in sorted(vlinfo):
            if isstring:
                value = spssaux._smartquote(value)
            else:
                if value == int(value):
                    value = int(value)
            label = spssaux._smartquote(label)
            vls.append("%s %s" % (value, label))
        return "VALUE LABELS " + varname + "\n   " + "\n   ".join(vls) + "."
        
    def report(self, labelsyntax):
        # write report
        
        if not labelsyntax:
            print _("""No value labels were generated.""")
            return
        
        if len(self.labelvars) > 1:
            labelvars = self.labelvars
        else:
            labelvars = len(self.varstolabel) * [self.labelvars][0]
        spss.StartProcedure("Generate Value Labels", "STATSVALLBLSFROMDATA")            
        cells = [[labelvars[i], 
            spss.CellText.Number(len(self.conflicts[vname]), spss.FormatSpec.Count), 
            #spss.CellText.Number(self.duplabels[vname], spss.FormatSpec.Count)]\
            spss.CellText.Number(len(self.duplabels[vname]), spss.FormatSpec.Count)]\
            for i,vname in enumerate(self.varstolabel)]
        caption = []
        if self.syntax:
            caption.append(_("""Generated label syntax: %s""" % self.syntax))
        if self.execute:
            caption.append(_("""Generated label syntax was applied"""))
        else:
            caption.append(_("""Generated label syntax was not applied"""))
        caption.append(_("""A conflict means that different labels would be applied to the same value."""))
        caption.append(_("""A duplicate means that the same label was used for different values."""))
            
        tbl = spss.BasePivotTable(_("""Value Label Generation"""), "VALLBLSFROMDATA",
            caption="\n".join(caption))
        tbl.SimplePivotTable(rowdim= _("""Variable"""), rowlabels=self.varstolabel, 
            collabels=[_("""Label Source"""), _("""Label Conflicts"""), _("""Duplicate Labels""")],
            cells=cells)
        spss.EndProcedure()
        
    def truncate(self, name, maxlength):
        """Return a name truncated to no more than maxlength BYTES.
        
        name is the candidate string
        maxlength is the maximum byte count allowed.  It must be a positive integer.
        
        If name is a (code page) string, truncation is straightforward.  If it is Unicode utf-8,
        the utf-8 byte representation must be used to figure this out but still truncate on a character
        boundary."""
        
        if name is None:
            return None
        if not self.unicodemode:
            name =  name[:maxlength]
        else:
            newname = []
            nnlen = 0
            
            # In Unicode mode, length must be calculated in terms of utf-8 bytes
            for c in name:
                c8 = self.ec(c)[0]   # one character in utf-8
                nnlen += len(c8)
                if nnlen <= maxlength:
                    newname.append(c)
                else:
                    break
            name = "".join(newname)
        if name[-1] == "_":
            name = name[:-1]
        return name    

def mkrandomname(prefix="D", sav=True):
    res = prefix + str(random.uniform(.01,1.0))
    if sav:
        res = res + ".sav"
    return res

    def __init__(self):
        self.wdsname = mkrandomname("D", sav=False)
        
    def getsav(self, filespec, delete=True):
        """Open sav file and return all contents
        
        filespec is the file path
        filespec is deleted after the contents are read unless delete==False"""
     
        item = self.wdsname
        spss.Submit(r"""get file="%(filespec)s".
DATASET NAME %(item)s.
DATASET ACTIVATE %(item)s.""" % locals())
        contents = spssdata.Spssdata(names=False).fetchall()
        spss.Submit("""DATASET CLOSE %(item)s.
        NEW FILE.""" % locals())
        if delete:
            os.remove(filespec)
        return contents

    """return list of syntax specs
    
    root is the root of the name for the left hand variable
    setn is the set number
    setvars is the list of variables in the set
    data is the table of unstandardized canconical coefficients
        - one column per canonical correlation
    ndims can trim the number of correlations used."""
    
    syntax = []
    nvars = len(setvars)
    ncor = len(data[0])
    if not ndims is None:
        ncor = min(ncor, ndims)
    newnames = set()
    for i in range(ncor):
        cname = root + "_set" + str(setn) + "_" + str(i+1)
        newnames.add(cname)
        if len(cname) > 64:
            raise ValueError(_("The specified root name is too long: %s") % root)
        s = ["COMPUTE " + cname + " = "]
        for j in range(nvars):
            s.append(str(data[j][i]) + " * " + setvars[j])
        syntax.append(s[0] + " + ".join(s[1:]))
        syntax[i] = "\n".join(textwrap.wrap(syntax[i])) +"."

    return "\n".join(syntax), newnames



        
class NonProcPivotTable(object):
    """Accumulate an object that can be turned into a basic pivot table once a procedure state can be established"""
    
    def __init__(self, omssubtype, outlinetitle="", tabletitle="", caption="", rowdim="", coldim="", columnlabels=[],
                 procname="Messages"):
        """omssubtype is the OMS table subtype.
        caption is the table caption.
        tabletitle is the table title.
        columnlabels is a sequence of column labels.
        If columnlabels is empty, this is treated as a one-column table, and the rowlabels are used as the values with
        the label column hidden
        
        procname is the procedure name.  It must not be translated."""
        
        attributesFromDict(locals())
        self.rowlabels = []
        self.columnvalues = []
        self.rowcount = 0

    def addrow(self, rowlabel=None, cvalues=None):
        """Append a row labelled rowlabel to the table and set value(s) from cvalues.
        
        rowlabel is a label for the stub.
        cvalues is a sequence of values with the same number of values are there are columns in the table."""

        if cvalues is None:
            cvalues = []
        self.rowcount += 1
        if rowlabel is None:
            self.rowlabels.append(str(self.rowcount))
        else:
            self.rowlabels.append(rowlabel)
        self.columnvalues.extend(cvalues)
        
    def generate(self):
        """Produce the table assuming that a procedure state is now in effect if it has any rows."""
        
        privateproc = False
        if self.rowcount > 0:
            try:
                table = spss.BasePivotTable(self.tabletitle, self.omssubtype)
            except:
                StartProcedure(_("Create dummy variables"), self.procname)
                privateproc = True
                table = spss.BasePivotTable(self.tabletitle, self.omssubtype)
            if self.caption:
                table.Caption(self.caption)
            if self.columnlabels != []:
                table.SimplePivotTable(self.rowdim, self.rowlabels, self.coldim, self.columnlabels, self.columnvalues)
            else:
                table.Append(spss.Dimension.Place.row,"rowdim",hideName=True,hideLabels=True)
                table.Append(spss.Dimension.Place.column,"coldim",hideName=True,hideLabels=True)
                colcat = spss.CellText.String("Message")
                for r in self.rowlabels:
                    cellr = spss.CellText.String(r)
                    table[(cellr, colcat)] = cellr
            if privateproc:
                spss.EndProcedure()
                
def attributesFromDict(d):
    """build self attributes from a dictionary d."""
    self = d.pop('self')
    for name, value in d.iteritems():
        setattr(self, name, value)

def StartProcedure(procname, omsid):
    """Start a procedure
    
    procname is the name that will appear in the Viewer outline.  It may be translated
    omsid is the OMS procedure identifier and should not be translated.
    
    Statistics versions prior to 19 support only a single term used for both purposes.
    For those versions, the omsid will be use for the procedure name.
    
    While the spss.StartProcedure function accepts the one argument, this function
    requires both."""
    
    try:
        spss.StartProcedure(procname, omsid)
    except TypeError:  #older version
        spss.StartProcedure(omsid)

class FileHandles(object):
    """manage and replace file handles in filespecs.
    
    For versions prior to 18, it will always be as if there are no handles defined as the necessary
    api is new in that version, but path separators will still be rationalized.
    """
    
    def __init__(self):
        """Get currently defined handles"""
        
        # If the api is available, make dictionary with handles in lower case and paths in canonical form, i.e.,
        # with the os-specific separator and no trailing separator
        # path separators are forced to the os setting
        if os.path.sep == "\\":
            ps = r"\\"
        else:
            ps = "/"
        try:
            self.fhdict = dict([(h.lower(), (re.sub(r"[\\/]", ps, spec.rstrip("\\/")), encoding))\
                for h, spec, encoding in spss.GetFileHandles()])
        except:
            self.fhdict = {}  # the api will fail prior to v 18
    
    def resolve(self, filespec):
        """Return filespec with file handle, if any, resolved to a regular filespec
        
        filespec is a file specification that may or may not start with a handle.
        The returned value will have os-specific path separators whether or not it
        contains a handle"""
        
        parts = re.split(r"[\\/]", filespec)
        # try to substitute the first part as if it is a handle
        parts[0] = self.fhdict.get(parts[0].lower(), (parts[0],))[0]
        return os.path.sep.join(parts)
    
    def getdef(self, handle):
        """Return duple of handle definition and encoding or None duple if not a handle
        
        handle is a possible file handle
        The return is (handle definition, encoding) or a None duple if this is not a known handle"""
        
        return self.fhdict.get(handle.lower(), (None, None))
    
    def createHandle(self, handle, spec, encoding=None):
        """Create a file handle and update the handle list accordingly
        
        handle is the name of the handle
        spec is the location specification, i.e., the /NAME value
        encoding optionally specifies the encoding according to the valid values in the FILE HANDLE syntax."""
        
        spec = re.sub(r"[\\/]", re.escape(os.path.sep), spec)   # clean up path separator
        cmd = """FILE HANDLE %(handle)s /NAME="%(spec)s" """ % locals()
        # Note the use of double quotes around the encoding name as there are some encodings that
        # contain a single quote in the name
        if encoding:
            cmd += ' /ENCODING="' + encoding + '"'
        spss.Submit(cmd)
        self.fhdict[handle.lower()] = (spec, encoding)
        
def Run(args):
    """Execute the STATS VALLBLS FROMDATA extension command"""

    args = args[args.keys()[0]]

    oobj = Syntax([
        Template("VARIABLES", subc="",  ktype="varname", var="variables", islist=True),
        Template("VARPATTERN", subc="",  ktype="literal", var="varpattern", islist=False),
        Template("LBLVARS", subc="",  ktype="varname", var="lblvars", islist=True),
        Template("LBLPATTERN", subc="",  ktype="literal", var="lblpattern", islist=False),

        Template("VARSPERPASS", subc="OPTIONS", ktype="int", var="varsperpass"),

        Template("SYNTAX", subc="OUTPUT", ktype="literal", var="syntax"),
        Template("EXECUTE", subc="OUTPUT", ktype="bool", var="execute"),
        
        Template("HELP", subc="", ktype="bool")])
    
    #enable localization
    global _
    try:
        _("---")
    except:
        def _(msg):
            return msg
    # A HELP subcommand overrides all else
    if args.has_key("HELP"):
        #print helptext
        helper()
    else:
        processcmd(oobj, args, dolabels)

def helper():
    """open html help in default browser window
    
    The location is computed from the current module name"""
    
    import webbrowser, os.path
    
    path = os.path.splitext(__file__)[0]
    helpspec = "file://" + path + os.path.sep + \
         "markdown.html"
    
    # webbrowser.open seems not to work well
    browser = webbrowser.get()
    if not browser.open_new(helpspec):
        print("Help file not found:" + helpspec)
try:    #override
    from extension import helper
except:
    pass        