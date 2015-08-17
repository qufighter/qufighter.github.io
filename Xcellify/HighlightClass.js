// quick way to provide a list of classes to highlight 

var highlightClasses = {
  applyCss: function(classColors){
    var matches, m, ml, css;
    for( var cl in classColors ){
      matches = document.querySelectorAll('.'+cl);
      for( m=0,ml=matches.length; m<ml; m++ ){
        for( css in classColors[cl] ){
          matches[m].style[css] = classColors[cl][css];
        }
      }
    }
  },
  showKey: function(elm, classColors){
    var css, keyElm;
    for( var cl in classColors ){
      keyElm = Cr.elm('span',{},[Cr.txt(cl)],elm);
      for( css in classColors[cl] ){
        keyElm.style[css] = classColors[cl][css];
      }
    }
  }
}