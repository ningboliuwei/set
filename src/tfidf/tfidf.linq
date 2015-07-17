<Query Kind="Program">
  <Reference>&lt;RuntimeDirectory&gt;\System.Net.Http.dll</Reference>
  <Reference>&lt;RuntimeDirectory&gt;\System.Web.dll</Reference>
  <Reference>&lt;ProgramFilesX86&gt;\Microsoft ASP.NET\ASP.NET MVC 4\Assemblies\System.Web.Http.dll</Reference>
  <Namespace>System.Net.Http</Namespace>
  <Namespace>System.Net.Http.Headers</Namespace>
</Query>

void Main()
{
    // see:https://en.wikipedia.org/wiki/Tf%E2%80%93idf        
    var doc1 = new List<string>(){
        "this",
        "is",
        "a",
        "a",
        "sample"
    };
    
    var doc2 = new List<string>(){
        "this",
        //"this",
        "is",
        "a",
        //"a",
        "sample"
        //"another",
        //"another",
        //"example",
        //"example",
        //"example"
    };
    
    var doc3 = new List<string>(){
        "this",
        "is",
        "a",
        "a",
        "sample"
    };
    
    var doc4 = new List<string>(){
        "this",
        "is",
        "a",
        "two",
        "sample"
    };
    
    var doc5 = new List<string>(){
        "this",
        "is",
        "ccc",
        "sss",
        "sample"
    };
    
    var doc6 = new List<string>(){
        "this",
        "is",
        "ccc",
        "aasss",
        "sample"
    };
    
    var doc7 = new List<string>(){
        "this",
        "is",
        "a",
        "two",
        "sample"
    };
    

    
    var docs = new Dictionary<string,List<string>>{
        {"doc1",doc1},
        {"doc2",doc2},
        {"doc3",doc3},
        {"doc4",doc4},
        {"doc5",doc5},
        {"doc6",doc6},
        {"doc7",doc7}
    };
    
    var fdIdfs = docs.ToFDIDF(2,
        f=>f,
        (N,d)=>1+Math.Log(N*1.0/d,10)
    );    
    
    fdIdfs.Rank().Dump();
}

// Define other methods and classes here
public static class Extension{
    public static List<Tuple<string,List<double>>> ToFDIDF(this Dictionary<string,List<string>> docs,int N,Func<int,int> tfFun,Func<int,int,double> idfFun){
        // caculate term frequency
        var tfs = 
        docs.AsParallel()
            .Select(d=>d.Value.GroupBy(w=>w).Select(g=>new {doc = d.Key,word = g.Key,value=tfFun(g.Count())}))
            .Dump();
         
        // calculate invers documents frequency
        var idfs = 
        tfs.Aggregate((a,b)=>a.Union(b))
           .GroupBy(a=>a.word)
           .Select(g=> new {
            word = g.Key,
            value = idfFun(N,g.Count())
           })
           .Dump();
           
        // calculate td-idfs
        var tfidfs = new List<Tuple<string,List<double>>>();
        foreach(var tf in tfs){
            var tfidf = new List<double>();
            var tfDict = tf.ToDictionary(t=>t.word,t=>t.value);
            foreach(var idf in idfs){
                if(tfDict.ContainsKey(idf.word)){
                    tfidf.Add(tfDict[idf.word]*idf.value);
                }else{
                    tfidf.Add(0);
                }
            }
            tfidfs.Add(Tuple.Create(tf.First().doc,tfidf));
        }
        return tfidfs;
    }
    public static double Cos(this List<double> V1, List<double> V2){
        int N = ((V2.Count < V1.Count)?V2.Count : V1.Count);
        double dot  = 0.0d;
        double mag1 = 0.0d;
        double mag2 = 0.0d;
        for (int n = 0; n < N; n++){
            dot += V1[n] * V2[n];
            mag1 += Math.Pow(V1[n], 2);
            mag2 += Math.Pow(V2[n], 2);
        }
        return dot / (Math.Sqrt(mag1) * Math.Sqrt(mag2));
    }
    public static IEnumerable<Tuple<string,double>> Rank(this List<Tuple<string,List<double>>> list){
        var item1 = list[0];
        var query = list.Select(item=>Tuple.Create(item.Item1,item.Item2.Cos(item1.Item2))).OrderBy(v=>v.Item2);
        foreach(var i in query){
            yield return i;
        }
    }
}