from pySW import SW
import psutil, time, shutil
from pyDOE import *
import pandas as pd
#
partName    = r'Box.SLDPRT';

if "SLDWORKS.exe" in (p.name() for p in psutil.process_iter()) == False:
    print('starting SLDWORKS')
    SW.startSW();
    time.sleep(10);

SW.connectToSW()

SW.openPrt(psutil.os.getcwd()+'\\main\\'+partName);

var = SW.getGlobalVariables();

"""
FUNCTIONALITY TO BE ADDED IN THE UPCOMING VERSION TO GET A LIST OF VARIABLES
RATHER THAN A DICTIONARY
"""
variables = [];
for key in var:
    variables.append(key);


design = lhs(len(var.keys()), samples=5)
means = [50, 50, 50]
stdvs = [10, 5, 15]
from scipy.stats.distributions import norm
for i in range(len(var.keys())):
    design[:, i] = norm(loc=means[i], scale=stdvs[i]).ppf(design[:, i])


analysisDir = psutil.os.getcwd()+'\\analysis';
if psutil.os.listdir(psutil.os.getcwd())[0] == 'analysis':
    shutil.rmtree(analysisDir);
    psutil.os.mkdir(analysisDir);
else:
    psutil.os.mkdir(analysisDir);


for i in range(len(design)):
    psutil.os.mkdir(analysisDir+'\\'+str(i));
    for j in range(len(variables)):
        SW.modifyGlobalVar(variables[j],round(design[i][j]),'mm');
        SW.updatePrt();
        SW.saveAssy(analysisDir+'\\'+str(i), str(i), '.SLDPRT');


df = pd.DataFrame(design, columns = variables)
df.to_csv(psutil.os.getcwd()+'\\params.csv')
