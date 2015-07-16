<Query Kind="Program">
  <Reference>&lt;RuntimeDirectory&gt;\System.Net.Http.dll</Reference>
  <Reference>&lt;RuntimeDirectory&gt;\System.Web.dll</Reference>
  <Reference>&lt;ProgramFilesX86&gt;\Microsoft ASP.NET\ASP.NET MVC 4\Assemblies\System.Web.Http.dll</Reference>
  <Namespace>System.Net.Http</Namespace>
  <Namespace>System.Net.Http.Headers</Namespace>
</Query>

void Main()
{
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
    
    var docs = new Dictionary<string,List<string>>{
        {"doc1",doc1},
        {"doc2",doc2}
    };
    
    var fdIdfs = docs.ToFDIDF(2,
        f=>f,
        (N,d)=>1+Math.Log(N*1.0/d,10)
    );
    
    fdIdfs["doc1"].Cos(fdIdfs["doc2"]).Dump();
}

// Define other methods and classes here
public static class Extension{
    public static Dictionary<string,List<double>> ToFDIDF(this Dictionary<string,List<string>> docs,int N,Func<int,int> tfFun,Func<int,int,double> idfFun){
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
        var tfidfs = new Dictionary<string,List<double>>();
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
            tfidfs.Add(tf.First().doc,tfidf);
        }
        return tfidfs;
    }
    public static double Cos(this List<double> V1, List<double> V2){
        int N = 0;
        N = ((V2.Count < V1.Count)?V2.Count : V1.Count);
        double dot = 0.0d;
        double mag1 = 0.0d;
        double mag2 = 0.0d;
        for (int n = 0; n < N; n++)
        {
            dot += V1[n] * V2[n];
            mag1 += Math.Pow(V1[n], 2);
            mag2 += Math.Pow(V2[n], 2);
        }
    
        return dot / (Math.Sqrt(mag1) * Math.Sqrt(mag2));
    }
}
