import os
rootDir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
reportDir = rootDir + "\\reports\\xml-report\\"
print(type(max([os.path.join(reportDir,d) for d in os.listdir(reportDir)], key=os.path.getmtime)))
a = str(max([os.path.join(reportDir,d) for d in os.listdir(reportDir)], key=os.path.getmtime))

          