function onOpen() {
  var ui = DocumentApp.getUi();

  ui.createMenu('Custom Options')
      .addItem('Convert To HTML', 'convert_to_html')
      .addToUi();
}

/**
 * Convert the current document to html
 */
function convert_to_html() {
  // Get the document to which this script is bound.
  var doc = DocumentApp.getActiveDocument();
  
  var doc_converter = new Doc_To_HTML();
  doc_converter.convert_and_save_to_file(doc);
  
}


function Doc_To_HTML(){
  var self = this;

  self.special_chars
  
  /**
   * Convert a document to html
   * @param document <Document>
   * @return html <string>
   */
  self.convert = function(document){
    var body = document.getBody();
    
    html = self.build_tags(body);
    return html;
  }
  
  /**
   * Convert a document to html and save to a new file on drive
   * @param document <Document>
   * @return html <string>
   */
  self.convert_and_save_to_file = function(document){
    var html = self.convert(document);
    var filename = document.getName() + ' [Converted to HMTL]'
    
    DriveApp.createFile(filename, html, MimeType.HTML);
    DocumentApp.getUi().alert("A new HTML file '" + filename + "' was created.");  
  }
  
  /**
   * Build the html tags from the document
   * @param body <Body>
   * @return html <string>
   */
  self.build_tags = function (body){
    if (typeof body.getNumChildren == "undefined"){
      //There are no children
      return self.find_links_in_text(body);
    }
    var html_text = ""
    var number_of_children = body.getNumChildren();
    
    for(var child_index = 0; child_index < number_of_children; child_index++){
      var child = body.getChild(child_index);
      var type = child.getType();
      
      var tags = ["",""];
      
      switch(type){
        case DocumentApp.ElementType.PARAGRAPH:
          var heading = child.getHeading()
          tags = self.get_heading_tags(heading);
          break;
        case DocumentApp.ElementType.LIST_ITEM:
          tags = self.get_listing_tags(body, child_index, number_of_children)
          break;
        case DocumentApp.ElementType.TEXT:
          tags = ["",""]
          break;
      }
      html_text += tags[0] + self.build_tags(child) + tags[1];
    }  
    return html_text;
  }
  
  /**
   * Takes a heading and figures out which heading tag to apply to it
   * @param heading <Element>
   * @return opening_and_closing_tags <array>
   */
  self.get_heading_tags = function (heading){
    switch(heading){
      case DocumentApp.ParagraphHeading.HEADING1:
        return ["<h1>","</h1>"];
      case DocumentApp.ParagraphHeading.HEADING2:
        return ["<h2>","</h2>"];
      case DocumentApp.ParagraphHeading.HEADING3:
        return ["<h3>","</h3>"];
      default:
        return ["<p>","</p>"]
    }
  }
  
  /**
   * Tasks a list and generates tags for that list. If first element of the list prepends <ul> if last element appends </ul>
   * @param body <Body or Paragraph>
   * @param child_index <int>
   * @param number_of_children <int>
   * @return tags <array> - opening and closing
   */
  self.get_listing_tags = function (body, child_index, number_of_children){
    tags = ["<li>","</li>"]
    if(child_index == 0 || body.getChild(child_index - 1).getType() != DocumentApp.ElementType.LIST_ITEM){
      tags[0] = "<ul>" + tags[0];
    }
    if(child_index + 1 == number_of_children || body.getChild(child_index + 1).getType() != DocumentApp.ElementType.LIST_ITEM){
      tags[1] += "</ul>";
    }
    return tags
  }
  
  /**
   * Finds links within text and adds anchor tags within the text
   * @param text <Element>
   * @return html <string>
   */
  self.find_links_in_text = function (text){
    var text_length = text.getText().length;
    var chars = text.getText();
    var html = "";
    var link_text = undefined;
    
    for(var i = 0; i < text_length; i++){
      var url = text.getLinkUrl(i);
      var char = self.convert_special_characters(chars[i]);
      if(url) {
        if(link_text && link_text.url != url){
          //back to back urls, but this one is different
          html += "<a href='" + link_text.url + "'>" + link_text.anchor_text + "</a>";
          link_text = { url: url, anchor_text: char };
        }
        else if(!link_text) {
          link_text = { url: url, anchor_text: char };
        }
        else {
          link_text.anchor_text += char;
        }
      }
      else{
        if(link_text){
          html += "<a href='" + link_text.url + "'>" + link_text.anchor_text + "</a>";
        }
        link_text = undefined;
        html += char
      }
    }   
    if(link_text){
      html += "<a href='" + link_text.url + "'>" + link_text.anchor_text + "</a>";
    }
    
    return html;
  }
  
  /**
   * Finds special characters and converts them to HTML entities
   * @param char <string>
   * @return <srting>
   */
  self.convert_special_characters = function(char){
    var output = HtmlService.createHtmlOutput(char);
    return output.getContent();
  }
}
