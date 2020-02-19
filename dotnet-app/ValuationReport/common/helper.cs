using System.Collections.Generic;
using System.IO;
using System.Linq;

class Helper{
    public List<string> splitString( string toSplit , string delimiter)
    {
        var returnList = toSplit.Split(delimiter).ToList();
        return returnList;
    }
    public string checkFileOrDirectory (string path){
        string isFileOrDirectory = null;
        return isFileOrDirectory = File.Exists(path) ? "file" : "directory";
        //if (File.Exists(path))
        //{
        //    pathis   = "file";
        //}
        //else if(Directory.Exists(path))
        //{
        //     pathis   = "directory";
        //}
        //return pathis;
    }
    public string[] getFileNames( string directoryPath , string fileExtension)
    {
        var files = Directory.GetFiles(directoryPath, fileExtension);        
        return files;
    }

    public void writeToTextFile(string path , string toBeWrittten)
    {
        if(!File.Exists(path))
        {
            using (StreamWriter sw = File.CreateText(path))
            {
                sw.WriteLine(toBeWrittten);
            }

        }
        using(StreamWriter sw = File.AppendText(path))
        {
            sw.WriteLine(toBeWrittten);
        }

    }


}