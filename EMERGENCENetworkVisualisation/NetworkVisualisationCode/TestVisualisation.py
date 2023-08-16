from pyvis.network import Network
from constantly import ValueConstant, Values
from textwrap3 import wrap
import networkx as nx, matplotlib.pyplot as plt, xml.etree.ElementTree as ET, openpyxl, random  
import streamlit as st, streamlit.components.v1 as components, pandas as pd, networkx as nx

class Specialty(Values):
   """
   Constants representing various specialties across the EMERGENCE network. 
   """
   COMPUTER_SCIENCE = ValueConstant("Computer Science")
   ROBOTICS = ValueConstant("Robotics")
   HEALTHCARE = ValueConstant("Healthcare")
   DESIGN_ENGINEERING_INNOVATION = ValueConstant("Design Engineering and Innovation")
   ELECTRONIC_ENGINEERING = ValueConstant("Electronic Engineering")
   SOCIAL_CARE = ValueConstant("Social Care")
   PHYSIOTHERAPY = ValueConstant("Physiotherapy")
   INCLUDE_CENTRAL_NODE = ValueConstant("include-central-node")
   INCLUDE_LEGEND_NODES = ValueConstant("include-legend-nodes")

class Colour(Values):
   """
   Constants representing the various colours of the nodes within the network visualisation
   """
   RED = ValueConstant("red")
   BLUE = ValueConstant("blue")
   GREEN = ValueConstant("green")
   YELLOW = ValueConstant("yellow")
   ORANGE = ValueConstant("orange")
   MAGENTA = ValueConstant("magenta")
   PURPLE = ValueConstant("purple")

class Misc(Values):
   """
   Constants representing miscellaneous values used within the program.  
   """
   EXCEL_FILE = ValueConstant("C:\\Users\\maxim\\OneDrive\\Documents\\EMERGENCENetworkVisualisation\\NetworkVisualisationData\\EMERGENCECollatedData.xlsx")
   ZERO = ValueConstant(0)
   ONE = ValueConstant(1)
   TWO = ValueConstant(2)
   FIVE = ValueConstant(5)
   SIX = ValueConstant(6)
   SEVEN = ValueConstant(7)
   EIGHT = ValueConstant(8)
   NINE = ValueConstant(9)
   TEN = ValueConstant(10)
   ELEVEN = ValueConstant(11)
   TWELVE = ValueConstant(12)
   THIRTEEN = ValueConstant(13)
   LEGEND_TITLE = ValueConstant("Legend Node: ")
   ONE_HUNDRED = ValueConstant(100)
   LEGEND_X_VALUE = ValueConstant(-1500)
   LEGEND_Y_VALUE = ValueConstant(-1250)
   LEGEND_NODE_SIZE = ValueConstant(75)
   LEGEND_SHAPE = ValueConstant("box")
   LEGEND_WIDTH_CONSTRAINT = ValueConstant(150)
   LEGEND_FONT_SIZE = ValueConstant(20)
   CENTRAL_NODE_ID = ValueConstant("Central Node")
   CENTRAL_NODE_LABEL = ValueConstant("Emergence Network")
   CENRAL_NODE_COLOUR = ValueConstant("black")
   HTML_FILE_NAME = ValueConstant("network_visualisation.html")
   NAME = ValueConstant("Name: ")
   INSTITUTION = ValueConstant("\nInstitution: ")
   JOB_TITLE = ValueConstant("\nJob Title: ")
   SPECIALTY = ValueConstant("\nSpecialty: ")
   LOCATION_SECTION = ValueConstant("\nLocation: ")
   RESEARCH_THEMES = ValueConstant("\nResearch themes of interest: ")
   TRUE = ValueConstant(True)
   FALSE = ValueConstant(False)
   NEWLINE = ValueConstant("\n")
   VALUE = ValueConstant("value")
   ID = ValueConstant("id")
   TITLE = ValueConstant("title")
   ALPHABET = ValueConstant("alphabet")
   GROUP = ValueConstant("group")
   LABEL = ValueConstant("label")
   SIZE = ValueConstant("size")
   FIXED = ValueConstant("fixed")
   X = ValueConstant("x")
   Y = ValueConstant("y")
   PHYSICS = ValueConstant("physics")
   SHAPE = ValueConstant("shape")
   WIDTH_CONSTRAINT = ValueConstant("widthConstraint")
   FONT = ValueConstant("font")
   SPECIALTY_WORD = ValueConstant("Specialty")
   INSTITUTION_WORD = ValueConstant("institution")
   SIZE = ValueConstant("size")
   LOCATION = ValueConstant("location")
   PX = ValueConstant('{fpixels}px')
   HEIGHT = ValueConstant("1080px")
   WIDTH = ValueConstant("100%")
   BGCOLOUR = ValueConstant("#fffff")

def setMemberColour(specialty):
   """
   Returns a specific node-colour based on the member's specialty.  
   """
   match specialty:
      #Matches the specialty to each case returning a specific colour
      case Specialty.COMPUTER_SCIENCE.value:
         #All computer scientists have a red node colour
         return Colour.RED.value
      case Specialty.ROBOTICS.value:
         #All roboticists have a blue node colour 
         return Colour.BLUE.value
      case Specialty.DESIGN_ENGINEERING_INNOVATION.value:
         #All design engineers have a green node colour
         return Colour.GREEN.value
      case Specialty.HEALTHCARE.value:
         #All Heathcare professionals/students have a yellow node colour
         return Colour.YELLOW.value
      case Specialty.PHYSIOTHERAPY.value:
         #All Physiotherapists have an orange node colour
         return Colour.ORANGE.value
      case Specialty.ELECTRONIC_ENGINEERING.value:
         #All Electronic Engineers have a magenta node colour
         return Colour.MAGENTA.value
      case Specialty.SOCIAL_CARE.value:
         #All Social Care workers have a purple node colour 
         return Colour.PURPLE.value

def format_motivation_text(motivationText, charLimit):
   """
   formats the passed in 'motivationText' to abide by the line character limit
   charLimit.
   """
   splitText = wrap(motivationText, charLimit)
   #formats 'motivationText' in to charLimit character blocks
   #each of these blocks are stored in the list splitText
   formattedMessage=Misc.NEWLINE.value + Misc.NEWLINE.value
   #Creates a two-line gap between the title and the text
   for i in range(len(splitText)):
      #Iterates through the 'splitText' array and concatenates charLimit
      #characters before a newline
      formattedMessage += splitText[i] + Misc.NEWLINE.value
   #returns a formattedMessage once this is complete
   return formattedMessage


#List containing the names of all the members in the network 
member_names = []

#Reference to the excel file containing the member details 
member_workbook = openpyxl.load_workbook(Misc.EXCEL_FILE.value)

sheet = member_workbook.worksheets[0]

print(sheet)

#Initialises a networkX graph for the legend nodes
nx_graph = nx.Graph()

#Stores the number of actual nodes & the number of nodes in the legend
num_actual_nodes = sheet.max_row-Misc.ONE.value
num_legend_nodes = Misc.SEVEN.value

#The Step value represents the spacing between each of the legend nodes
step = Misc.ONE_HUNDRED.value
#The x value represents the x position of the legend 
x = Misc.LEGEND_X_VALUE.value
#The y value represents the y position of the legend
y = Misc.LEGEND_Y_VALUE.value
#A list of all of the labels for the legend nodes
legend_labels = [Specialty.ROBOTICS.value,Specialty.HEALTHCARE.value,Specialty.COMPUTER_SCIENCE.value, Specialty.DESIGN_ENGINEERING_INNOVATION.value,
                 Specialty.ELECTRONIC_ENGINEERING.value, Specialty.SOCIAL_CARE.value, Specialty.PHYSIOTHERAPY.value]

#A list of tuples containing all of the legend nodes
legend_nodes = [
    (
        num_actual_nodes + legend_node, #Node ID set to the sum of the current numbher of nodes  
                                                                # + the index of the legend node
        {
            Misc.GROUP.value: legend_node, #Adds the new legend node to the group
            Misc.LABEL.value: legend_labels[legend_node],#Indexes the legend_labels list to return the appropriate label for the node
            Misc.SIZE.value: Misc.LEGEND_NODE_SIZE.value,#Defines the size of the legend nodes
            Misc.FIXED.value: Misc.TRUE.value,#Sets the legend position to fixed
            Misc.PHYSICS.value: Misc.FALSE.value, #Sets the legend to have no physics mechanics
            Misc.X.value: x, #Sets the x-value of the legend
            Misc.Y.value: Misc.PX.value.format(fpixels=y+legend_node*step), #Sets the y-value of the legend
            Misc.SHAPE.value: Misc.LEGEND_SHAPE.value, #Sets the shape of the legend nodes
            Misc.WIDTH_CONSTRAINT.value: Misc.LEGEND_WIDTH_CONSTRAINT.value, #Sets the width contraints for the legend
            Misc.FONT.value: {Misc.SIZE.value: Misc.LEGEND_FONT_SIZE.value}, #Sets the font size for the legend
            Misc.SPECIALTY_WORD.value: Specialty.INCLUDE_LEGEND_NODES.value, #Sets the speciality attribute for the legend 'include-legend-nodes'
            Misc.INSTITUTION_WORD.value: Specialty.INCLUDE_LEGEND_NODES.value,#Sets the institution attribute for the legend 'include-legend-nodes'
            Misc.ALPHABET.value: Specialty.INCLUDE_LEGEND_NODES.value, #Sets the alphabet attribute for the legend 'include-legend-nodes'
            Misc.TITLE.value: Misc.LEGEND_TITLE.value + legend_labels[legend_node], #Sets the title attribute for the legend 'include-legend-nodes'
            Misc.LOCATION.value: Specialty.INCLUDE_LEGEND_NODES.value
        }
    )
    for legend_node in range(num_legend_nodes)#Creates 'num_legend_nodes' legend nodes and specifies an attribute for all of them
]

nx_graph.add_nodes_from(legend_nodes) #Adds all of the legend nodes to the graph

nx_graph.add_node(Misc.CENTRAL_NODE_ID.value, label=Misc.CENTRAL_NODE_LABEL.value, title=Misc.CENTRAL_NODE_ID.value, physics=Misc.TRUE.value, color=Misc.CENRAL_NODE_COLOUR.value, location= Specialty.INCLUDE_CENTRAL_NODE.value, institution=Specialty.INCLUDE_CENTRAL_NODE.value, Specialty=Specialty.INCLUDE_CENTRAL_NODE.value, alphabet=Specialty.INCLUDE_CENTRAL_NODE.value, x=1000, y=1000)    
#Adds in the central node of the network visualisation which is labeled "EMERGENCE_Network"

for row_index in range(Misc.TWO.value, sheet.max_row+Misc.ONE.value):
     #Iterates through all of the rows in the excel document
          #if the row is not the title row 
          member_name = str(sheet.cell(row = row_index, column= Misc.SIX.value).value)
          print(member_name)
          #The member_name variable is extracted from the fifth column of the sheet
          member_institution = str(sheet.cell(row = row_index, column= Misc.SEVEN.value).value)
          #The member_institution variable is extracted from the sixth column of the sheet
          member_location = str(sheet.cell(row = row_index, column= Misc.EIGHT.value).value)
          #The member_location variable is extracted from the seventh column of the sheet
          member_job_title = str(sheet.cell(row = row_index, column= Misc.NINE.value).value)
          #The member_job_title variable is extracted from the eighth column of the sheet
          member_specialty = str(sheet.cell(row = row_index, column= Misc.TEN.value).value)
          #The member_specialty variable is extracted from the ninth column of the sheet
          member_colour = setMemberColour(member_specialty)
          #The member_name variable is assigned the result of the setMemberColour function
          member_network_motivation = format_motivation_text(str(sheet.cell(row = row_index, column= Misc.THIRTEEN.value).value), Misc.ONE_HUNDRED.value)
          #The member_network_motivation variable is assigned the result of the format_motivation_text function
          member_alphabet = member_name[Misc.ZERO.value]
          #The member_specialty variable is extracted from the ninth column of the sheet
          member_title = Misc.NAME.value + member_name + Misc.INSTITUTION.value + member_institution + Misc.JOB_TITLE.value + member_job_title + Misc.SPECIALTY.value + member_specialty + Misc.LOCATION_SECTION.value + member_location + Misc.RESEARCH_THEMES.value + member_network_motivation
          #The member_title variable displays all of the member attributes when the node is hovered over
          nx_graph.add_node(member_name, label=member_name, title=member_title, physics=Misc.TRUE.value,  institution=member_institution, Specialty=member_specialty, alphabet=member_alphabet, color=member_colour, location=member_location)
          #Creates a new node for each of the EMERGENCE members with all of the relevant attributes
          member_names.append(member_name)
          #Appends each member to the member_names array

for i in range(Misc.ZERO.value, len(member_names)):
   #Iterates through the member_names array
   nx_graph.add_edge(Misc.CENTRAL_NODE_ID.value, member_names[i]) #Creates an edge between the central node and all the other nodes

G = Network(notebook=True, filter_menu=Misc.TRUE.value, height=Misc.HEIGHT.value, width=Misc.WIDTH.value, bgcolor=Misc.BGCOLOUR.value, cdn_resources='remote')
#Defines the PyVis graph which takes in the initial networkX graph as an input.
G.from_nx(nx_graph) 
#Derives a PyVis graph from the previous networkX graph

#Sets the title of the StreamLit Website
st.title('Network Graph Visualisation of the EMERGENCE Network')

neighbour_map = G.get_adj_list() #Gets the list of all neighbours to a node

for node in G.nodes:
   #Iterates through all of the nodes in the graph
  node[Misc.VALUE.value] = len(neighbour_map[node[Misc.ID.value]])#Sets the node size based on the number of neighbours

# Save and read graph as HTML file (on Streamlit Sharing)

G.save_graph('test_visualisation.html')
HtmlFile = open('test_visualisation.html','r',encoding='utf-8')

components.html(HtmlFile.read(), width=1920, height=1080)