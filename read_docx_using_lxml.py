from lxml import etree
import zipfile

ooXMLns = {
   'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
}

def get_comments(docxFileName):
   docxZip = zipfile.ZipFile(docxFileName)
   #docxZip.extractall()
   commentsXML = docxZip.read('word/comments.xml')
   et = etree.XML(commentsXML)
   comments = et.xpath('//w:comment', namespaces = ooXMLns)
   for c in comments:
      # attributes:
      comment_string="\nComment Id : {}, Comment Author : {}, Comment Date : {}, Comment : {}".format(c.xpath('@w:id', namespaces = ooXMLns)[0],c.xpath('@w:author', namespaces = ooXMLns), (c.xpath('@w:date', namespaces = ooXMLns)),(c.xpath('string(.)', namespaces = ooXMLns)) )
      print(comment_string)
      
get_comments("test.docx")