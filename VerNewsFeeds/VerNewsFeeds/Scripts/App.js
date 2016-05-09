'use strict';
var posts = []; // guardara posts
var resultados=[];

function getQueryStringParameter(requestedParameter){
    
    var param = document.URL.split("?")[1].split("&");
    
    var strParams = "";
    for(var i=0; i<param.length; i++){
        var actual = param[i].split("=");
        if(actual[0]== requestedParameter){
            return actual[1];
        }
    }
    
}

// aqui voy a decodificar
var appWebUrl = decodeURIComponent(getQueryStringParameter("SPAppWebUrl"));
var hostWebUrl = decodeURIComponent(getQueryStringParameter("SPHostUrl"));
// contexto de aplicacion
var clientContext = new SP.clientContext.get_current();
// contexto del tenant sharepoint
var hostContext = new SP.AppContextSite(clientContext,hostWebUrl);
// necesito formdigets para escribir
var formDigest = "";
var getFormDigest = function(){
    // contextinfo me devuelve toda la información de la aplicación
    
    $.ajax({
       url:appWebUrl+"/_api/contextinfo",
       type:"POST",
       contentType:"application/json;odata=verbose",
       headers:{
           'accept':'application/json;odata=verbose'
       },
       success:function(data){
         formDigest = data.d.GetContextWebInformation.formDigestValue;  
       },
       error:function(res){
           alert(res.responseText);
       },
       async:false // sera sincrono
    });
}(); // autoejecutable para obtener FORMDIGEST

// COMO VAMOS A UTILIZAR LAS CARACTERISTICAS SOCIALES
// ES NECESARIO LA INFORMACIÓN DEL AUTOR
var getActorInfo=function(cuenta){
    var actor = "";
    // si tienes alguna tuberia
    if(cuenta.indexOf("|")>0)
    cuenta = cuenta.split("|")[2];
    // el actor sera el nombre de quien va acceder a todos los post
    $.ajax({
        url:appWebUrl+"/_api/social.feed/actor(item='"+cuenta+"')",
        headers:{
            "accept":"application/json;odata=verbose"
        },
        success:function(data){
          actor = data.d.FollowableItemActor;  
        },
        error:function(res){
          alert("Error:"+res.responseText);  
        },
        async:false
    });
    return actor;
}

var getSiteFeed = function(){
    // obtener feeds
    var feed;
    $.ajax({
        //@ indica que es un parametro variable
        url:appWebUrl+"/_api/social.feed/actor(item=@v)/feed?@v='"+hostWebUrl+"/newsfeed.aspx'",
        headers:{
            "accept":"application/json;odata=verbose"
        },
        success:function(data){
            feed= getFeeds(data);
        },
        error:function(xhr){
            alert("Error:"+xhr.responseText);
        }
    });
    return feed;
}

function getFeeds(data){
    posts = data.d.SocialFeed.Threads.results.reverse();
    var query = "(ContentTypeId:0x01FD4FB0210AB50249908EAA47E6BD3CFE8B* OR ContentTypeId:0x01FD59A0DF25F1E14AB882D2C87D4874CF84* OR ContentTypeId:0x012002* OR ContentTypeId:0x0107* OR WebTemplate=COMMUNITY)";
    // si quieramos filtar por site, podriamos usar owstaxIdMetadataAllTagsInfo:" + Utilities.projectSiteCode
    
    
    
    var keywordQuery = new Microsoft.SharePoint.Client.Search.Query.keywordQuery(clientContext);
    keywordQuery.set_queryText(query);
    var lista = keywordQuery.get_sortList();
    lista.add("LastModifiedTime",Microsoft.SharePoint.Client.Search.Query.SortDirection.Ascending);
    keywordQuery.set_enableSorting(true);
    var executor = new Microsoft.SharePoint.Client.Search.Query.SearchExecutor(clientContext);
    var results = executor.executeQuery(keywordQuery);
    
    clientContext.executeQueryAsync(function(){
        // voy acceder al resultado
        // guardamos los resultados de la busqueda
        resultados = results.m_value.ResultTables[0].ResultRows;
        // updateDisplay recarga o repinta todos los resultados con los post
        updateDisplay();
         
    },function(e){
        alert("Error");
    });
    
}

var updateDisplay = function(){
    var postb;
    var post;
    var contenido="";
    while(post.length!=0 || resultados.length!=0){
        if(post.length==0){
            // lenght-1 para obtener de atras hacia adelante
            post = resultados[resultados.lenght-1];
            var autor = getActorInfo(postb.PostAuthor);
            // voy a recibir Html
            contenido += addToFeed(autor,postb.FullPostBody,postb.Created);
            resultados.pop();
        } else if(resultados.length==0){
            post = post[posts.length-1].Actors.results[post.Authorindex];
            var autor = posts[posts.length-1].Actors.results[post.Authorindex];
            contenido += addToFeed(autor,post.FullPostBody,new Date(postb.Created);
            posts.pop(); // elimino el ultimo elemento del array
        }else{
            postb = resultados[resultados.length-1];
            post = posts[posts.length-1].RootPost;
            // si la fecha del post es mayor a la fecha del
            if(new Date(post.CreatedTime)>postb.Created){
                var autor = posts[posts.length-1].Actors.results[post]
                contenido += addToFeed(autor,
                post.FullPostBody,
                new Date(post.CreatedTime));
                posts.pop();
            }else{
                postb = resultados[resultados.length-1];
                var autor = getActorInfo(postb.PostAuthor);
                contenido += addToFeed(autor,postb.FullPostBody,postb.Created);
                resultados.pop();
            }
        }
    }
    contenido+="</ul>";
    $("#Posts").html(contenido); // agrego el contenido HTML
}

function addToFeed(autor,texto,fecha){
    var contenido ="<li>"+autor.name +"<br/>"+texto+"<br />"+fecha+"</li>";
    return contenido;
}

function sendPost(){
    var contenido =$("#mensaje").val();
    $.ajax({
        url:appWebUrl+"/_api/social.feed/actor(item=@v)/feed/post?@v="+hostWebUrl+"/newsfeed.aspx",
        type:"POST",
        data:JSON.stringify({
            "restCreationData":{
                "__metadata":{
                    "type":"SP.Social.SocialRestPostCreationData"      
                },
                "ID":null,
                "creationData":{"__metadata":{
                    
                }}
            },
        }),
        headers:{
          "accept":"application/json;odata=verbose",
          "content-type":"application/json;odata=verbose",
          "X-RequestDigest":formDigest  
        },
        success:getFeeds,
        error:function(){
            alert("Error");
        }
    });
}

$(document).ready(function()){
    getSiteFeeds();
    $("#Post").click(sendPost());
})
    



