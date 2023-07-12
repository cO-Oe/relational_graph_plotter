import openpyxl
import networkx as nx
import matplotlib.pyplot as plt

class Node:
  def __init__(self, color, _id, text):
    self.color = color
    self._id = _id
    self.text = text

class Edge:
  def __init__(self, line_type, src, des):
    self.line_type = line_type
    self.src = src
    self.des = des
  
  def __hash__(self):
    return hash((self.line_type, self.src, self.des))

  def __eq__(self, other):
    return ((self.src == other.src and self.line_type == other.line_type and self.des == other.des) or (self.des == other.src and self.line_type == other.line_type and self.src == other.des))

def read_excel():
  nodes = []
  edges = set([])

  workbook = openpyxl.load_workbook(r'./DATA-Biocomputing.xlsx')
  sheet = workbook.active

  for row in sheet.iter_rows(min_row=2):
    nodes.append(Node(row[1].value, row[2].value, row[3].value))
    for i in range(4, len(row), 2):
      if row[i+1].value == None:
        break
      elif row[i+1].value == 'linkReverse':
        edges.add(Edge('link', row[i].value, row[2].value))
      else:
        edges.add(Edge(row[i+1].value, row[2].value, row[i].value))

  return nodes, edges

def draw_graph(nodes, edges):
  labels = {}
  for node in nodes:
    labels[node._id] = node.text

  red_nodes = [node._id for node in nodes if node.color == 'red']
  blue_nodes = [node._id for node in nodes if node.color == 'blue']

  link = [(edge.src, edge.des) for edge in edges if edge.line_type == 'link']
  synonymLink = [(edge.src, edge.des) for edge in edges if edge.line_type != 'link']
  synonymLink = [tuple(sorted(t)) for t in synonymLink]
  synonymLink = list(set(synonymLink))

  G = nx.DiGraph(format='png', directed=True)
  G.add_nodes_from(red_nodes+blue_nodes)
  G.add_edges_from(link+synonymLink)

  pos = nx.planar_layout(G)
  label_pos = {}

  for p in pos:
    label_pos[p] = [pos[p][0]+0.04, pos[p][1]]

  nx.draw_networkx_nodes(G, pos, nodelist=red_nodes, node_color='r', node_size=300)
  nx.draw_networkx_nodes(G, pos, nodelist=blue_nodes, node_color='b', node_size=300)
  nx.draw_networkx_labels(G, label_pos, labels=labels, font_size=10, font_weight=1, horizontalalignment='left')

  ax = plt.gca()
  for edge in link:
    ax.annotate("",
                xy=pos[edge[0]],
                xytext=pos[edge[1]],
                arrowprops=dict(arrowstyle="<|-", color="b",
                                shrinkA=7, shrinkB=7,
                                connectionstyle="arc3,rad=-0.3",
                ),
    )
  for edge in synonymLink:
    ax.annotate("",
                xy=pos[edge[0]],
                xytext=pos[edge[1]],
                arrowprops=dict(arrowstyle="-", color="0.4",
                                shrinkA=7, shrinkB=7,
                                linestyle="dotted",
                                connectionstyle="arc3,rad=-0.3",
                ),
    )

  x_vals, y_vals = zip(*label_pos.values())
  x_max, x_min = max(x_vals), min(x_vals)
  y_max, y_min = max(y_vals), min(y_vals)
  x_margin = (x_max - x_min) * 0.3
  y_margin = (y_max - y_min) * 0.15
  plt.xlim(x_min - x_margin, x_max + x_margin)
  plt.ylim(y_min - y_margin, y_max + y_margin)
  
  ax.axis('off')
  plt.get_current_fig_manager().full_screen_toggle()
  plt.show()

def main():
  nodes, edges = read_excel()

  draw_graph(nodes, edges)

if __name__ == '__main__':
  main()