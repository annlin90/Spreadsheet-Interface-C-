using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ComponentModel;
using System.Xml;
using System.IO;


namespace SpreadsheetEngine
{
    public class Class1
    {
    }
}

namespace CptS321
{
    public abstract class Cell : INotifyPropertyChanged //abstract class
    {
        public event PropertyChangedEventHandler PropertyChanged;
        private readonly int RowIndex = 0;
        private readonly int ColumnIndex = 0;
        protected string Value = " ";
        protected string Text = " ";

        public Cell(int RowIndex2, int ColumnIndex2) //Constructor for cell 
        {
            RowIndex = RowIndex2;
            ColumnIndex = ColumnIndex2;
        }

        public string Values
        {
            get  //getter for value property
            {
                return Value;
            }
        }

        public int Rows
        {
            get //getter for RowIndex
            {
                return RowIndex;
            }
        }

        public int Columns
        {
            get //getter for ColumnIndex
            {
                return ColumnIndex;
            }
        }

        public string Texter
        {
            get //getter for Text property
            {
                return Text;
            }

            set //setter for Text Property
            {
                if (Text != value) //if text is being changed
                {
                    Text = value; //then text becomes updated
                    PropertyChanged(this, new PropertyChangedEventArgs("Text")); //fire the PropertyChanged event
                }
                else return; //if text == value
            }
        }
    }

    public class Spreadsheet
    {
        public event PropertyChangedEventHandler CellPropertyChanged;
        public Cell[,] Arrays; //Array of Cells

        private class CellHelper : Cell //Inherit from Cell Class
        {
            public CellHelper(int Rows, int Columns) : base(Rows, Columns)
            { //calls Cell.Cell(Rows, Columns);
            }

            public void setValue(string value)
            {
                Value = value; //sets value of cell
            }
        }

        public int RowCount
        {
            get //gets the # of rows
            {
                return Arrays.GetLength(0);
            }
        }

        public int ColumnCount
        {
            get //gets the # of columns
            {
                return Arrays.GetLength(1);
            }
        }

        public Cell GetCell(int Rows, int Columns)
        {
            return Arrays[Rows, Columns]; //returns cell at specific location
        }

        public Spreadsheet(int Rows, int Columns)
        {
            Arrays = new CellHelper[Rows, Columns];
            //spreadsheet constructor that takes in a # of rows and columns

            for (int i = 0; i < Rows; i++)
            {
                for (int j = 0; j < Columns; j++)
                {
                    CellHelper Temps = new CellHelper(i, j);
                    Temps.PropertyChanged += OnPropertyChanged;
                    Arrays[i, j] = Temps;
                }
            }
        }

        public Cell getLetterLocation(string places)
        {
            int col = places[0] - 'A';
            int row = Int32.Parse(places.Substring(1)) - 1; //gets cell through string location

            return GetCell(row, col);

        }

        public void OnPropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == "Text") //PropertyChanged events, lets them know when properties for each cell has changed
            {
                CellHelper tmpCell = sender as CellHelper;
            }
            else
            {
                CellPropertyChanged(sender, new PropertyChangedEventArgs("Value")); //fires if changed
            }
            contentsExpress(sender as Cell);
        }

        public void contentsExpress(Cell cell)
        {
            CellHelper cellOne = cell as CellHelper; //create new cells
            Cell cellTwo;
            string formula = "";

            if (cellOne.Texter[0] == '=')  //if there's a formula
            {
                formula = cellOne.Texter.Substring(1); //get whatever's after the '=' sign
                cellTwo = getLetterLocation(formula); //get the location that's after the '=' sign
                cellOne.setValue(cellTwo.Texter); //then set the value of the cell that has the formula to the value of the cell it's asking

            }

            else if (string.IsNullOrEmpty(cellOne.Texter)) //if cell is empty set it to empty
            {
                cellOne.setValue("");
            }
            else
            {
                cellOne.setValue(cellOne.Texter); //set value 

            }
            if (CellPropertyChanged != null)
            {
                CellPropertyChanged(cell, new PropertyChangedEventArgs("Value")); //fire if changed
            }
        }
        //saveToXML function, translates data recieved from individual cell classes to XML format
        public void XMLSave(FileStream files) //maybe this needs to return an XML document...
        {
            var XML = new XmlDocument();
            XmlWriterSettings settings = new XmlWriterSettings();
            settings.Indent = true;
            settings.OmitXmlDeclaration = true;
            settings.NewLineOnAttributes = true;


            using (XmlWriter Writing = XmlWriter.Create(files, settings))
            {
                Writing.WriteStartElement("Spreadsheet"); //This will create the "root"
                foreach (Cell cells in Arrays)
                {
                    if (cells.Texter != "")    //it has been edited
                    {
                        string cellName = Convert.ToString(Convert.ToChar(cells.Columns + 65)) + Convert.ToString(cells.Columns + 1);
                        Writing.WriteStartElement("cell");
                        Writing.WriteAttributeString("name", cellName);

                        Writing.WriteStartElement("text");
                        Writing.WriteString(cells.Texter);
                        Writing.WriteEndElement();          //for cText

                        Writing.WriteEndElement();          //for cell
                    }
                }

                Writing.WriteEndElement(); // for spreadsheet
          }
        }
        //load the spreadsheet from XML file and call AddCell to insert them into the spreadsheet
        public void XMLLoad(FileStream XMLFile)
        {
            string C_Name = "";
            //int bgcolor = 0;
            string Temps = "";

            clearAll();
             
            XmlReaderSettings Sett = new XmlReaderSettings();
            Sett.DtdProcessing = DtdProcessing.Parse;
            XmlReader User = XmlReader.Create(XMLFile, Sett);
            while (User.Read())
            {
                switch (User.NodeType)
                {
                    case XmlNodeType.Element:
                        if (User.Name == "cell")
                        {
                            User.MoveToNextAttribute(); //loading from xml onto spreadsheet
                            C_Name = User.Value;
                            break;
                        }
                         
                        else if (User.Name == "text")
                        {
                            User.Read();
                            Temps = User.Value;
                        }
                        break;
                    case XmlNodeType.EndElement:
                        if (User.Name == "cell")
                        {
                            addCell(C_Name, Temps);
                            C_Name = "";
                            Temps = "";
                             
                        }
                        break;
                    default:
                        break;
                }
            }
        }

        //adds a cell to the spreadsheet array for use in the loadFromXML function
        public void addCell(string first, string Second)
        {
            string Columns1 = first.Substring(0, 1);      
            string Rows1 = first.Substring(1);      
            int Column2 = Convert.ToChar(Columns1) - 65;                 
            int Rows2 = Convert.ToInt32(Rows1) - 1;

            if (Second != "")
                Arrays[Rows2, Column2].Texter = Second;
            
        }

        //function to clear the spreadsheet before load function is called
        private void clearAll()
        
        {
            for (int x = 0; x < RowCount; x++)
            {
                for (int y = 0; y < ColumnCount; y++)
                {
                    Arrays[x, y].Texter = ""; 
                }
            }
        }
    }


  public class ExpTree
        {
           private static Dictionary<string, double> ExpDict = new Dictionary<string, double>();
           private Node treeRoots;
            
          public Node ConstructTree(string Exp) //Constructs the expression tree
            {
            int i = 0;
            var num = 0.0;
            opNode currents = null;

            for (i = Exp.Length - 1; i >= 0; i--) {
                 switch (Exp[i]) {//creats node for each operator
                      case '/':
                      currents = new opNode(Exp[i], ConstructTree(Exp.Substring(0, i)), ConstructTree(Exp.Substring(i + 1)));
                      return currents;  
                      case '*':
                      currents = new opNode(Exp[i], ConstructTree(Exp.Substring(0, i)), ConstructTree(Exp.Substring(i + 1)));
                        return currents;
                       case '-':
                      currents = new opNode(Exp[i], ConstructTree(Exp.Substring(0, i)), ConstructTree(Exp.Substring(i + 1)));
                      return currents;
                        case '+':
                      currents = new opNode(Exp[i], ConstructTree(Exp.Substring(0, i)), ConstructTree(Exp.Substring(i + 1)));
                        return currents;
                    }
                }

            if (double.TryParse(Exp, out num))   //if can be parsed to double
                return new numNode(num); //return numNode
            else
                return new varNode(Exp); //else the varNode      
            }

        public abstract class Node {
            public abstract double Eval();
        }

        public class varNode : Node {
            private string Names;
            public varNode(string name) //creats varNode
            {
                this.Names = name;
                ExpDict[Names] = 0;
            }

            public override double Eval() {
                return ExpDict[Names];
            }
        }

        public class opNode : Node
        {
            private char Nodes;
            private Node Lefts;
            private Node Rights; //node values

            public opNode(char op, Node left, Node right) {
                this.Lefts = left; //create opNode
                this.Rights = right;
                this.Nodes = op;
            }

            public override double Eval() {
                switch (Nodes) //Evaluates the expression tree
                {
                  case '/':
                    return this.Lefts.Eval() / this.Rights.Eval();
                    case '*':
                      return this.Lefts.Eval() * this.Rights.Eval();
                    case '-':
                     return this.Lefts.Eval() - this.Rights.Eval();
                    case '+':
                        return this.Lefts.Eval() + this.Rights.Eval();
                }
                return 0;
            }
        }

        public class numNode : Node
        {
            private double valueing;
            public numNode(double value) {
                this.valueing = value; //makes numNode with value
            }

            public override double Eval()
            {
                return valueing;
            }
        }

        public void SetVar(string varName, double varValue)
        {
            if (ExpDict.ContainsKey(varName)) //Sets the specified variable variable within the ExpTree variables dictionary
                ExpDict[varName] = varValue;
            else
                ExpDict.Add(varName, varValue);
        }

        public ExpTree(string Exp)
        {
            treeRoots = ConstructTree(Exp); //calls to make the tree with the input expression
        }

        public double Eval()
        {
            if (treeRoots != null) //evaluates expression to a double value
                return treeRoots.Eval();

            else
                return 0.0;
        }
  
            public void Clear()
           {
                ExpDict.Clear(); //clears the dictionary
            }
        }
   
}


