"""
# *****************************************************************************
# *                                                                           *
# *    Copyright (C) 2020  Kalyan Inamdar, kalyaninamdar@protonmail.com       *
# *                                                                           *
# *    This library is free software; you can redistribute it and/or          *
# *    modify it under the terms of the GNU Lesser General Public             *
# *    License as published by the Free Software Foundation; either           *
# *    version 2 of the License, or (at your option) any later version        *
# *                                                                           *
# *    This library is distributed in the hope that it will be useful,        *
# *    but WITHOUT ANY WARRANTY; without even the implied warranty of         *
# *    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the          *
# *    GNU Lesser General Public License for more details.                    *
# *                                                                           *
# *    You should have received a copy of the GNU General Public License      *
# *    along with this program.  If not, see http://www.gnu.org/licenses/.    *
# *                                                                           *
# *****************************************************************************
"""
import subprocess as sb
import win32com.client
import pythoncom
import os
#
class commSW:
    def __init__(self):
        pass;
    #
    def startSW(self, *args):
        #                                                                     #
        # Function to start Solidworks from Python.                           #
        #                                                                     #
        # Accepts an optional argument: the year of version of Solidworks.    #
        #                                                                     #
        # If you have only one version of Solidworks on your computer, you do #
        # not need to provide this input.                                     #
        #                                                                     #
        # Example: If you have Solidworks 2019 and Solidworks 2020 on your    #
        # system and you want to start Solidworks 2020 the function you call  #
        # should look like this: startSW(2020)                                #
        #                                                                     #
        if not args:
            SW_PROCESS_NAME = r'C:/Program Files/SOLIDWORKS Corp/SOLIDWORKS/SLDWORKS.exe';
            sb.Popen(SW_PROCESS_NAME);
        else:
            year= int(args[0][-1]);
            SW_PROCESS_NAME = "SldWorks.Application.%d" % (20+(year-2));
            win32com.client.Dispatch(SW_PROCESS_NAME);
    #
    def shutSW(self):
        #                                                                     #
        # Function to close Solidworks from Python.                           #
        # Does not accept any input.                                          #
        #                                                                     #
        sb.call('Taskkill /IM SLDWORKS.exe /F');
    #
    def connectToSW(self):
        #                                                                     #
        # Function to establish a connection to Solidworks from Python.       #
        # Does not accept any input.                                          #
        #                                                                     #
        global swcom
        swcom = win32com.client.Dispatch("SLDWORKS.Application");
    #
    def openAssy(self, prtNameInp):
        #                                                                     #
        # Function to open an assembly document in Solidworks from Python.    #
        #                                                                     #
        # Accepts one input as the filename with the path if the working      #
        # directory of your script and the directory in which the assembly    #
        # file is saved are different.                                        #
        #                                                                     #
        self.prtNameInn = prtNameInp;
        self.prtNameInn = self.prtNameInn.replace('\\','/');
        #
        if os.path.basename(self.prtNameInn).split('.')[-1].lower() == 'sldasm': 
            pass;
        else:
            self.prtNameInn+'.SLDASM'
        #
        openDoc     = swcom.OpenDoc6;
        arg1        = win32com.client.VARIANT(pythoncom.VT_BSTR, self.prtNameInn);
        arg2        = win32com.client.VARIANT(pythoncom.VT_I4, 2);
        arg3        = win32com.client.VARIANT(pythoncom.VT_I4, 0);
        arg5        = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 2);
        arg6        = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 128);
        #
        openDoc(arg1, arg2, arg3, "", arg5, arg6);
    #
    def openPrt(self, prtNameInp):
        #                                                                     #
        # Function to open an part document in Solidworks from Python.        #
        #                                                                     #
        # Accepts one input as the filename with the path if the working      #
        # directory of your script and the directory in which the part file   #
        # is saved are different.                                             #
        #                                                                     #
        self.prtNameInn = prtNameInp;
        self.prtNameInn = self.prtNameInn.replace('\\','/');
        #
        openDoc     = swcom.OpenDoc6;
        arg1        = win32com.client.VARIANT(pythoncom.VT_BSTR, self.prtNameInn);
        arg2        = win32com.client.VARIANT(pythoncom.VT_I4, 1);
        arg3        = win32com.client.VARIANT(pythoncom.VT_I4, 1);
        arg5        = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 2);
        arg6        = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 128);
        #
        openDoc(arg1, arg2, arg3, "", arg5, arg6);
    #
    def updatePrt(self):
        if 'model' in globals():
            pass;
        else:
            global model;
            model   = swcom.ActiveDoc;
        model.EditRebuild3;
    #
    def closePrt(self):
        swcom.CloseDoc(os.path.basename(self.prtName));
    #
    def saveAssy(self, directory, fileName, fileExtension):
        if 'model' in globals():
            pass;
        else:
            global model;
            model   = swcom.ActiveDoc;
        directory   = directory.replace('\\','/');
        comFileName = directory+'/'+fileName+'.'+fileExtension;
        arg         = win32com.client.VARIANT(pythoncom.VT_BSTR, comFileName);
        model.SaveAs3(arg, 0, 0);
    #
    def getGlobalVariables(self):
        #                                                                     #
        # Function to extract a set of global variables in a Solidworks       #
        # part/assembly file. The part/assembly is then automatically updated.#
        #                                                                     #
        # Does not accept any input.                                          #
        # Provides output in the form of a dictionary with Global Variables as#
        # the keys and values of the variables as values in the dictionary.   #
        #                                                                     #
        if 'model' in globals():
            pass;
        else:
            global model;
            model   = swcom.ActiveDoc;
        #
        if 'eqMgr' in globals():
            pass;
        else:
            global eqMgr;
            eqMgr = model.GetEquationMgr;
        #
        n = eqMgr.getCount;
        #
        data = {};
        #
        for i in range(n):
            if eqMgr.GlobalVariable(i) == True:
                data[eqMgr.Equation(i).split('"')[1]] = i
            #
        #
        if len(data.keys) == 0:
            raise KeyError("There are not any 'Global Variables' present in the currently active Solidworks document.");
        else:
            return data;
    #
    def modifyGlobalVar(self, variable, modifiedVal, unit):
        #                                                                     #
        # Function to modify a global variable or a set of global variables   #
        # in a Solidworks part/assembly file. The part/assembly is then       #
        # automatically updated.                                              #
        #                                                                     #
        # Accepts three inputs: variable name, modified value, and the unit   #
        # of the variable. The inputs can be string, integer and string       #
        # respectively or a list of variables, list of modified values and a  #
        # list of units of respective variables.                              #
        #                                                                     #
        # Note: In case you need to modify multiple dimensions using lists    #
        # the length of the lists must strictly be equal.                     #
        #                                                                     #
        if 'model' in globals():
            pass;
        else:
            global model;
            model   = swcom.ActiveDoc;
        #
        if 'eqMgr' in globals():
            pass;
        else:
            global eqMgr;
            eqMgr = model.GetEquationMgr;
        #
        data = self.getGlobalVariables();
        #
        if isinstance(variable, str) == True:
            eqMgr.Equation(data[variable], "\""+variable+"\" = "+str(modifiedVal)+unit+"");
        elif isinstance(variable, list) == True:
            if isinstance(modifiedVal, list) == True:
                if isinstance(unit, list) == True:
                    for i in range(len(variable)):
                        eqMgr.Equation(data[variable[i]], "\""+variable[i]+"\" = "+str(modifiedVal[i])+unit[i]+"");
                else:
                    raise TypeError("If a list of multiple variables is given, then lists of equal \n\
lengths should be given for 'modifiedVal' and 'unit' inputs.");
            else:
                raise TypeError("If a list of multiple variables is given, then lists of equal \n\
lengths should be given for 'modifiedVal' and 'unit' inputs.");
        else:
            raise TypeError("Incorrect input for the variables. Inputs can either be string, integer and string or lists containing variables, values and units.");
        #
        self.updatePrt();
    #
    def modifyLinkedVar(self, variable, modifiedVal, unit, *args):
        #                                                                     #
        # Function to modify a global variable/dimension or a set of          #
        # dimensions in a linked 'equations' file. The part/assembly is then  #
        # automatically updated.                                              #
        #                                                                     #
        # Accepts three inputs: variable name, modified value, and the unit   #
        # of the variable. The inputs can be string, integer and string       #
        # respectively or a list of variables, list of modified values and a  #
        # list of units of respective variables. Additionally the function    #
        # accepts one more optional argument, which is the complete path of   #
        # the equations file. If a path to the equations file is not provided #
        # then the function searches for a file named 'equations.txt' in the  #
        # working directory of the code.                                     #
        #                                                                     #
        # Note: In case you need to modify multiple dimensions using lists    #
        # the length of the lists must strictly be equal.                     #
        #                                                                     #
        #
        # Check the filename
        if len(args) == 0:
            file = 'equations.txt';
        else:
            file = args[0];
        #
        # READ FILE WITH ORIGINAL DIMENSIONS
        try:
            reader      = open(file, 'r');
        except IOError:
            raise IOError;
        finally:
            data = {};
            numLines    = len(reader.readlines());
            reader.close();
            reader      = open(file);
            lines       = reader.readlines();
            reader.close();
            for i in range(numLines):
                dim     = lines[i].split('"')[1];
                tempVal = lines[i].split(' ')[1];
                #
                val     = tempVal.replace(unit,'').replace('= ','').replace('\n','');
                data[dim] = val;
        #
        # MODIFY DIMENSIONS
        if isinstance(variable, list) == True:
            if isinstance(modifiedVal, list) == True:
                if isinstance(unit, list) == True:
                    for z in range(len(variable)):
                        data[variable[i]] = modifiedVal[i];
                else:
                    raise TypeError("If a list of multiple variables is given, then lists of equal \n\
lengths should be given for 'modifiedVal' and 'unit' inputs.");
            else:
                raise TypeError("If a list of multiple variables is given, then lists of equal \n\
lengths should be given for 'modifiedVal' and 'unit' inputs.");
        elif isinstance(variable, str) == True:
            data[variable] = modifiedVal;
        else:
            raise TypeError("The inputs types given.");
        #
        # WRITE FILE WITH MODIFIED DIMENSIONS
        writer      = open(file, 'w');
        for key, value in data.items():
            writer.write('"'+key+'"= '+str(value)+unit);
        writer.close();
        #
        self.updatePrt();
    #