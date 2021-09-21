using System;
using System.Linq;
using System.ComponentModel;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;

using SwissAcademic.Citavi;
using SwissAcademic.Citavi.Metadata;
using SwissAcademic.Citavi.Shell;
using SwissAcademic.Collections;

// Implementation of macro editor is preliminary and experimental.
// The Citavi object model is subject to change in future version.

public static class CitaviMacro
{
	public static void Main()
	{
		
		//Get the active project
		Project project = Program.ActiveProjectShell.Project;
		
		//Get the active ("primary") MainForm
		MainForm mainForm = Program.ActiveProjectShell.PrimaryMainForm;
		
		//Erstellt Liste mit zu betrachteten Gruppen
		List<Group> groups = project.Groups.ToList();
		groups.Sort();
		MessageBox.Show("Es erfolgt nun eine Abfrage welche der "+groups.Count+" Gruppen betrachtet werden sollen.");
		List<Group> selectedGroups = new List<Group>();
		foreach (Group groupp in groups){
			DialogResult dialogResult = MessageBox.Show("Soll Gruppe '"+groupp.Name+"' in Auswertung betrachtet werden?", "Frage", MessageBoxButtons.YesNo);
			if(dialogResult == DialogResult.Yes){
				selectedGroups.Add(groupp);
			}
		}
		//Analysiert Publikationen
		Dictionary<string, List<Reference>> groupAnz = new Dictionary<string, List<Reference>>();
	
		List<Reference> references = project.References.ToList();
		references.Sort();
		List<Reference> allReferencesList = new List<Reference>();
		foreach (Reference reference in references)
		{
			foreach(Group groupp in selectedGroups){				
				//Publikation der Gruppe zuordnen
				if (hasGroup(reference, groupp.Name)){
					//Publikation ist in ausgewählter Gruppe
					if(!allReferencesList.Contains(reference)){
						allReferencesList.Add(reference);
					}
					
					List<Reference> referenceList = null;
					if (groupAnz.TryGetValue(groupp.Name, out referenceList)) { 
						//Gruppe exisitert in Dictionary
						referenceList.Add(reference);
						groupAnz[groupp.Name] = referenceList;
					} else {
						//Gruppe exisitert nicht in Dictionary
						referenceList = new List<Reference>();
						referenceList.Add(reference);
						groupAnz.Add(groupp.Name,referenceList);
					}
					
				}
			}

		
		}
		//Debuging Message
		string message = "Liste der in den Gruppen gefundenen Anzahl von Arbeiten:\n";
		foreach (KeyValuePair<string, List<Reference>> kvp in groupAnz){
			message = message + "'"+kvp.Key+"': "+ kvp.Value.Count() +"\n";
			foreach (Reference r in kvp.Value){
				message = message + " - " + r.BibTeXKey;
			}
			message = message + " \n\n";
		}
		MessageBox.Show(message);
		
		//Erstelle Tex-Datei
		// create file, delete if already present
		StreamWriter sw = null;
		try 
		{
			//Hängt neuen Tect hinten an
			//sw = File.AppendText(createNewFile(mainForm));
			//Überschreibt bestehende Datei
			sw = File.CreateText(createNewFile(mainForm));
		}
		catch (Exception e)
		{
			DebugMacro.WriteLine("Error creating file: " + e.Message.ToString());
			return;
		}
		//Schreibt Header
		sw.WriteLine("{0}", getTexHeader());
		
		
		//Generiere Tabelle Publication * Group
		allReferencesList.Sort();
		selectedGroups.Sort();
		//Erstelle Header
		string str = "|l|";
		string groupHeader = "";
		foreach(Group groupp in selectedGroups){
			str = str +"c|";
			groupHeader = groupHeader + @" & \begin{sideways}"+"\n"+groupp.Name+"\n"+@"\end{sideways}"+"\n";
		}
		groupHeader = groupHeader + @"\\"+"\n"+@"\hline"+"\n";
		//Schreibe Header
		sw.WriteLine(@"\begin{tabular}{"+str+@"}"+"\n"+@"\hline" + "\n" + groupHeader);
		
		foreach(Reference reference in allReferencesList){
			str = reference.CitationKey.Replace("&","and") + @" \cite{" + reference.BibTeXKey + @"}";
			foreach(Group groupp in selectedGroups){
				if (hasGroup(reference, groupp.Name)) { 
					//Reference is in Group
						str = str + @" & $\bullet$ ";
					} else {
						str = str + @" & ";
				}
			}
			sw.WriteLine(str + @"\\" + "\n" + @"\hline");
			str = string.Empty;
		}
		//Schreibe Ende
		str = @"$ \sum $ ";
		foreach(Group groupp in selectedGroups){
			List<Reference> referenceList = null;
			if (groupAnz.TryGetValue(groupp.Name, out referenceList)) { 
				//Gruppe exisitert in Dictionary
				str = str + @" & " + referenceList.Count;
			} else {
				str = str + @" & 0 ";
			}
		}
		sw.WriteLine(str + @"\\" + "\n" + @"\hline");
		sw.WriteLine("\n" + @"\end{tabular}" + "\n");
//		\begin{center}
//		\begin{tabular}{ | c | c| c | } 
//		\hline
//			\begin{sideways} 
//					    Gruppe 1 
//				\end{sideways} & Gruppe2 & Gruppe3 \\ 
//		\hline
//		cell1 dummy text dummy text dummy text & $\bullet$ & cell6 \\ 
//		\hline
//		cell7 & cell8 & cell9 \\ 
//		\hline
//		\end{tabular}
//		\end{center}
		
		
		
		sw.WriteLine("\n \n \n \n");
		
		//Generiere gedrehte Tabelle Publication * Group
		allReferencesList.Sort();
		selectedGroups.Sort();
		//Erstelle Header
		str = "|l|";
		string pubHeader = "";
		//foreach(Group groupp in selectedGroups){
		foreach(Reference reference in allReferencesList){
			str = str +"c|";
			pubHeader = pubHeader + @" & \begin{sideways}"+"\n"+ reference.CitationKey.Replace("&","and") + @" \cite{" + reference.BibTeXKey + @"}" +"\n"+@"\end{sideways}"+"\n";
			//groupHeader = groupHeader + @" & \begin{sideways}"+"\n"+groupp.Name+"\n"+@"\end{sideways}"+"\n";
		}
		//Summenspallte
		pubHeader =  pubHeader + @" & "+"\n"+ @"$ \sum $ " +"\n";
		str = str +"c|";
		pubHeader = pubHeader + @"\\"+"\n"+@"\hline"+"\n";
		
		//Schreibe Header
		sw.WriteLine(@"\begin{tabular}{"+str+@"}"+"\n"+@"\hline" + "\n" + pubHeader);
		
		//foreach(Reference reference in allReferencesList){
		foreach(Group groupp in selectedGroups){
			//str = reference.CitationKey.Replace("&","and") + @" \cite{" + reference.BibTeXKey + @"}";
			str = groupp.Name;
			foreach(Reference reference in allReferencesList){
				if (hasGroup(reference, groupp.Name)) { 
					//Reference is in Group
						str = str + @" & $\bullet$ ";
					} else {
						str = str + @" & ";
				}
			}
			//Summe
			string sum;
			List<Reference> referenceList = null;
			if (groupAnz.TryGetValue(groupp.Name, out referenceList)) { 
				//Gruppe exisitert in Dictionary
				sum =  @" & " + referenceList.Count;
			} else {
				sum =  @" & 0 ";
			}
			
			sw.WriteLine(str + sum + @"\\" + "\n" + @"\hline");
			str = string.Empty;
		}
		//Schreibe Ende
		sw.WriteLine(str + @"\\" + "\n" + @"\hline");
		sw.WriteLine("\n" + @"\end{tabular}" + "\n");
		
		
		
		//Schreibt Footer
		sw.WriteLine("{0}", getTexFooter());
		
		sw.Close();
		MessageBox.Show("Done.");
	}
	
	public static bool hasGroup(Reference reference, String groupName){
		bool hasGroup = false;
		List<Group> groupsList = reference.Groups.ToList();
		groupsList.Sort();
		foreach (Group groupp in groupsList)
			{
				if (groupName.CompareTo(groupp.Name) == 0){
					hasGroup = true;
					break;
				}
			}
		return hasGroup;
	}
	
	public static string createNewFile(MainForm mainForm){
		string file = string.Empty;
		string initialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Personal);
		using (SaveFileDialog saveFileDialog = new SaveFileDialog())
		{
			saveFileDialog.Filter =  "Tex files (*.tex, *.txt)|*.tex;*.txt|All files (*.*)|*.*";
			saveFileDialog.InitialDirectory = initialDirectory;
			saveFileDialog.RestoreDirectory = true;
			saveFileDialog.Title = "Enter a file name for the Tex file.";	

			if (saveFileDialog.ShowDialog(mainForm) != DialogResult.OK) return file;
			file = saveFileDialog.FileName;
		}
		return file;
	}
	
	public static string groupsToStringList(List<Group> referenceGroups){
		string referenceGroupString = string.Empty;
		referenceGroups.Sort();
		foreach (Group referenceGroup in referenceGroups)
		{
			if (referenceGroupString == string.Empty){
				referenceGroupString = referenceGroup.Name;
			} else {
				referenceGroupString = referenceGroupString + "," + referenceGroup.Name;
			}
		}
		return referenceGroupString;
	}
	
public static string getRefernceYear(Reference reference){
	string year = reference.Year;
	year.Trim();
	if(year == ""){
		string[] bibtexList = reference.BibTeXKey.Split('.');
		year = bibtexList.Last();
	}
	if (year.Length >=5){
		year = year.Substring(year.Length-4, 4);
	}
	
	return year;
}
	
public static string getPlott(string groupName, List<Reference> references, string yearList){
	if (references.Count() == 0){
		return "";
	}
	references.Sort();
	//Ordne Referenzen Jahre zu
	Dictionary<string, List<Reference>> yearAnz = new Dictionary<string, List<Reference>>();
	foreach (Reference reference in references){
		List<Reference> referenceList = null;
		if (yearAnz.TryGetValue(getRefernceYear(reference), out referenceList)) { 
			//Gruppe exisitert in Dictionary
			referenceList.Add(reference);
			yearAnz[getRefernceYear(reference)] = referenceList;
		} else {
			//Gruppe exisitert nicht in Dictionary
			referenceList = new List<Reference>();
			referenceList.Add(reference);
			yearAnz.Add(getRefernceYear(reference),referenceList);
		}
	}
	//Baue Ausgabe zusammen
	string plott = @"
	\addplot coordinates {";
		
//	foreach (KeyValuePair<string, List<Reference>> kvp in yearAnz){
//		plott = plott + "(" + kvp.Key +","+ kvp.Value.Count +") \n";
//	}
	
	string[] years = yearList.Split(',');
	foreach (string year in years){
		List<Reference> refences;
		if (yearAnz.TryGetValue(year, out refences)) {
			plott = plott + "(" + year +","+ refences.Count +") \n";
		} else {
			//Jahr existert nicht in Liste
			//Füllt Graphen auf.
			plott = plott + "(" + year +",0) \n";
		}
	}
	
	
	plott = plott + @"
    };
	\addlegendentry{"+groupName+@"}";
	
//	\addplot[fill=green] coordinates {
//      (Unrenoviert,220.6)
//      (Fenster,219.26)
//      (Hlle,197.67)
//      (Bodenplatte,167.9)
//      (Heizung,40)
//    };";
	
	return plott;
}


public static string getTexHeader(){
string header =@"
\documentclass{article}
\usepackage{pgfplots}
\usepackage{rotating} %Für Tabelle Text rotieren
\pgfplotsset{
  compat=1.13,
}
\begin{document}";
	return header;
}


public static string getTexFooter(){
string footer =@"
\end{document}
";
	return footer;
	}
}