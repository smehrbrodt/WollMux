package de.muenchen.allg.itd51.wollmux.former.document.nodes;

import java.util.Collection;
import java.util.Iterator;

import de.muenchen.allg.itd51.wollmux.former.document.DocumentTree;
import de.muenchen.allg.itd51.wollmux.former.document.Visitor;

/**
 * Oberklasse für Knoten, die Nachfahren haben können (z,B, Absätze).
 * 
 * @author Matthias Benkmann (D-III-ITD 5.1)
 */
public class ContainerNode extends Node implements Container
{
  private Collection<Node> children;

  public ContainerNode(Collection<Node> children)
  {
    super();
    this.children = children;
  }

  public Iterator<Node> iterator()
  {
    return children.iterator();
  }

  public String toString()
  {
    return "CONTAINER";
  }

  public int getType()
  {
    return CONTAINER_TYPE;
  }

  public boolean visit(Visitor visit)
  {
    if (!visit.container(this, 0)) return false;

    Iterator<Node> iter = iterator();
    while (iter.hasNext())
    {
      if (!iter.next().visit(visit)) return false;
    }
    if (!visit.container(this, 1)) return false;
    return true;
  }
}