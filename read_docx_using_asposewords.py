import aspose.words as aw

def extract_content(startNode : aw.Node, endNode : aw.Node, isInclusive : bool):
  verify_parameter_nodes(startNode, endNode)
  nodes = []
  
  if (endNode.node_type == aw.NodeType.COMMENT_RANGE_END and isInclusive) :
    node = find_next_node(aw.NodeType.COMMENT, endNode.next_sibling)
    
    if (node != None) :
      endNode = node
  
  originalStartNode = startNode
  originalEndNode = endNode

  startNode = get_ancestor_in_body(startNode)
  endNode = get_ancestor_in_body(endNode)

  isExtracting = True
  isStartingNode = True

  currNode = startNode

  while (isExtracting) :
    cloneNode = currNode.clone(True)
    isEndingNode = currNode == endNode
    
    if (isStartingNode or isEndingNode) :
       if (isEndingNode) :
         process_marker(cloneNode, nodes, originalEndNode, currNode, isInclusive, False, not isStartingNode, False)
         isExtracting = False
       
       if (isStartingNode) :
         process_marker(cloneNode, nodes, originalStartNode, currNode, isInclusive, True, True, False)
         isStartingNode = False
    
    else :
      nodes.append(cloneNode)

    if (currNode.next_sibling == None and isExtracting) :
      nextSection = currNode.get_ancestor(aw.NodeType.SECTION).next_sibling.as_section()
      currNode = nextSection.body.first_child
    
    else :
      currNode = currNode.next_sibling
    
    if (isInclusive and originalEndNode == endNode and not originalEndNode.is_composite) :
      include_next_paragraph(endNode, nodes)
  return nodes

def verify_parameter_nodes(start_node: aw.Node, end_node: aw.Node):
  if start_node is None:
    raise ValueError("Start node cannot be None")
  
  if end_node is None:
    raise ValueError("End node cannot be None")
  
  if start_node.document != end_node.document:
    raise ValueError("Start node and end node must belong to the same document")
  
  if start_node.get_ancestor(aw.NodeType.BODY) is None or end_node.get_ancestor(aw.NodeType.BODY) is None:
    raise ValueError("Start node and end node must be a child or descendant of a body")
  
  start_section = start_node.get_ancestor(aw.NodeType.SECTION).as_section()
  end_section = end_node.get_ancestor(aw.NodeType.SECTION).as_section()
  start_index = start_section.parent_node.index_of(start_section)
  end_index = end_section.parent_node.index_of(end_section)
  
  if start_index == end_index:
    if (start_section.body.index_of(get_ancestor_in_body(start_node)) > end_section.body.index_of(get_ancestor_in_body(end_node))):
      raise ValueError("The end node must be after the start node in the body")
    
    elif start_index > end_index:
      raise ValueError("The section of end node must be after the section start node")

 
def find_next_node(node_type: aw.NodeType, from_node: aw.Node):
  if from_node is None or from_node.node_type == node_type:
    return from_node
  
  if from_node.is_composite:
    node = find_next_node(node_type, from_node.as_composite_node().first_child)
    if node is not None:
      return node

  return find_next_node(node_type, from_node.next_sibling)

 
def is_inline(node: aw.Node):
  return ((node.get_ancestor(aw.NodeType.PARAGRAPH) is not None or node.get_ancestor(aw.NodeType.TABLE) is not None) and not (node.node_type == aw.NodeType.PARAGRAPH or node.node_type == aw.NodeType.TABLE))

 
def process_marker(clone_node: aw.Node, nodes, node: aw.Node, block_level_ancestor: aw.Node,is_inclusive: bool, is_start_marker: bool, can_add: bool, force_add: bool):
  if node == block_level_ancestor:
    if can_add and is_inclusive:
      nodes.append(clone_node)
      return 
      assert clone_node.is_composite

  if node.node_type == aw.NodeType.FIELD_START: 
    if is_start_marker and not is_inclusive or not is_start_marker and is_inclusive:
      while node.next_sibling is not None and node.node_type != aw.NodeType.FIELD_END:
        node = node.next_sibling

  node_branch = fill_self_and_parents(node, block_level_ancestor)
  current_clone_node = clone_node

  for i in range(len(node_branch) - 1, -1):
    current_node = node_branch[i]
    node_index = current_node.parent_node.index_of(current_node)
    current_clone_node = current_clone_node.as_composite_node.child_nodes[node_index]
    remove_nodes_outside_of_range(current_clone_node, is_inclusive or (i > 0), is_start_marker)

  if can_add and (force_add or clone_node.as_composite_node().has_child_nodes):
    nodes.append(clone_node)

def remove_nodes_outside_of_range(marker_node: aw.Node, is_inclusive: bool, is_start_marker: bool):
  is_processing = True
  is_removing = is_start_marker
  next_node = marker_node.parent_node.first_child

  while is_processing and next_node is not None:

    current_node = next_node
    is_skip = False

    if current_node == marker_node:
      if is_start_marker:
        is_processing = False
        
        if is_inclusive:
          is_removing = False
        
        else:
          is_removing = True
          
          if is_inclusive:
            is_skip = True

    next_node = next_node.next_sibling
    if is_removing and not is_skip:
      current_node.remove()

 
def fill_self_and_parents(node: aw.Node, till_node: aw.Node):
  nodes = []
  current_node = node

  while current_node != till_node:
    nodes.append(current_node)
    current_node = current_node.parent_node

  return nodes

 
def include_next_paragraph(node: aw.Node, nodes):
  paragraph = find_next_node(aw.NodeType.PARAGRAPH, node.next_sibling).as_paragraph()
  if paragraph is not None:
    marker_node = paragraph.first_child if paragraph.has_child_nodes else paragraph
    root_node = get_ancestor_in_body(paragraph)
    process_marker(root_node.clone(True), nodes, marker_node, root_node,marker_node == paragraph, False, True, True)

 
def get_ancestor_in_body(start_node: aw.Node):
  while start_node.parent_node.node_type != aw.NodeType.BODY:
    start_node = start_node.parent_node
  return start_node

def generate_document(src_doc: aw.Document, nodes):
  dst_doc = aw.Document()
  dst_doc.first_section.body.remove_all_children()
  importer = aw.NodeImporter(src_doc, dst_doc,aw.ImportFormatMode.KEEP_SOURCE_FORMATTING)

  for node in nodes:
    import_node = importer.import_node(node, True)
    dst_doc.first_section.body.append_child(import_node)

  return dst_doc

def extract_comments(doc) :
    
  collectedComments = []
  comments = doc.get_child_nodes(aw.NodeType.COMMENT, True)
  for node in comments :
    comment = node.as_comment()
    comment_string= "Id: {0}, Author: {1}, DateTime: {2}, Comment: {3}".format(comment.id,comment.author,comment.date_time.strftime("%Y-%m-%d %H:%M:%S"),comment.to_string(aw.SaveFormat.TEXT))
    collectedComments.append(comment_string)     
  return collectedComments

def extract_paragraphs(doc) :
    
  collectedParagraphs = []
  paragraphs = doc.get_child_nodes(aw.NodeType.PARAGRAPH, True)

  for node in paragraphs :
    parah = node.as_paragraph()
    collectedParagraphs.append(parah.to_string(aw.SaveFormat.TEXT))
        
  return collectedParagraphs


#Code section for generating formatted_commented_document
# Load document.
doc = aw.Document("test.docx")
# Define starting and ending paragraphs.
startPara = doc.get_child(aw.NodeType.COMMENT_RANGE_START, 0, True).as_comment_range_start()
endPara = doc.get_child(aw.NodeType.COMMENT_RANGE_END , 0, True).as_comment_range_end()
# Extract the content between these paragraphs in the document. Include these markers in the extraction.
extractedNodes = extract_content(startPara, endPara, True)
doc.save("formatted_domment.docx")
#code section ends here


#Code section for getting document comments and paragraphs
extract_comments(doc)
extract_paragraphs(doc)
#code section ends here




