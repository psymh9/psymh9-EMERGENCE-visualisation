from git import Repo
from constantly import ValueConstant, Values
import os  

cwd = os.getcwd()
class Misc(Values):
   """
   Constants representing miscellaneous values used within the program.  
   """
   ORIGIN = ValueConstant('origin')
   ERROR_MESSAGE = ValueConstant('Some error occured while pushing the code')
   PATH_OF_GIT_REPO = ValueConstant(r'')
   COMMIT_MESSAGE = ValueConstant('New members added to the dataset')
   EXCEL_FILE = ValueConstant(cwd + "\\NetworkVisualisationData\\EMERGENCECollatedData.xlsm")

def git_push():
    #Checks for changes to the data file, if a new member has been added a commit is staged then pushed
    try:
        repo = Repo(Misc.PATH_OF_GIT_REPO.value)
        #Contains the repo object
        repo.git.add(Misc.EXCEL_FILE.value)
        #Stages the excel file containing all of the member data to git if there are any changes
        repo.index.commit(Misc.COMMIT_MESSAGE.value)
        #Commits the updated file with the commit message 'Update Network Visualisation'
        origin = repo.remote(name=Misc.ORIGIN.value)
        #Contains the origin branch of the repository
        origin.push()
        #pushes the changes up to the remote repository
    except:
        #If this process doesn't work print out an error message
        print(Misc.ERROR_MESSAGE.value)

git_push()