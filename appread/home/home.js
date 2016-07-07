(function(){
  'use strict';

  // The Office initialize function must be run each time a new page is loaded
  Office.initialize = function(reason){
    jQuery(document).ready(function(){});
    toArticleSchema()
  };

  window.toArticleSchema = function() {
    // Run a batch operation against the Word object model.
    Word.run(function (context) {
      // Create a proxy object for the document body.
      // The body object hasn't been set with any property values.
      var body = context.document.body;
      var bodyHTML = body.getHtml();

      var doc = {
        type: "doc",
        content: []
      }
      // Synchronize the document state by executing the queued commands,
      // and return a promise to indicate task completion.
      return context.sync().then(function () {
        $(bodyHTML.value).find('.Paragraph').each(function(index, element) {
          var _element = {}
          if ( element.hasAttribute('role') && element.getAttribute('role') === 'heading' ) {
            _element.type = 'heading1'
          } else {
            _element.type = 'paragraph'
          }
          _element.content = [
            {
              type: 'text',
              content: element.textContent,
              formats: []
            }
          ]
          doc.content.push(_element)
        });
        console.log(doc)
      });
    });
  }

})();
