/*
 * Dateiname: InsertionModelList.java
 * Projekt  : WollMux
 * Funktion : Verwaltet eine Liste von InsertionModels.
 * 
 * Copyright (c) 2008-2015 Landeshauptstadt München
 *
 * This program is free software: you can redistribute it and/or modify
 * it under the terms of the European Union Public Licence (EUPL),
 * version 1.0 (or any later version).
 *
 * This program is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * European Union Public Licence for more details.
 *
 * You should have received a copy of the European Union Public Licence
 * along with this program. If not, see
 * http://ec.europa.eu/idabc/en/document/7330
 *
 * Änderungshistorie:
 * Datum      | Wer | Änderungsgrund
 * -------------------------------------------------------------------
 * 06.09.2006 | BNK | Erstellung
 * -------------------------------------------------------------------
 *
 * @author Matthias Benkmann (D-III-ITD 5.1)
 * @version 1.0
 * 
 */
package de.muenchen.allg.itd51.wollmux.former.insertion;

import java.util.ArrayList;
import java.util.Collection;
import java.util.HashSet;
import java.util.Iterator;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.Vector;
import java.util.regex.Pattern;

import de.muenchen.allg.itd51.parser.ConfigThingy;
import de.muenchen.allg.itd51.parser.NodeNotFoundException;
import de.muenchen.allg.itd51.wollmux.L;
import de.muenchen.allg.itd51.wollmux.Logger;
import de.muenchen.allg.itd51.wollmux.UnknownIDException;
import de.muenchen.allg.itd51.wollmux.former.BroadcastListener;
import de.muenchen.allg.itd51.wollmux.former.BroadcastObjectSelection;
import de.muenchen.allg.itd51.wollmux.former.ComboboxMergeDescriptor;
import de.muenchen.allg.itd51.wollmux.former.FormularMax4kController;
import de.muenchen.allg.itd51.wollmux.former.IDManager;
import de.muenchen.allg.itd51.wollmux.former.control.FormControlModel;
import de.muenchen.allg.itd51.wollmux.former.function.FunctionSelectionAccess;

/**
 * Verwaltet eine Liste von InsertionModels
 * 
 * @author Matthias Benkmann (D-III-ITD 5.1)
 */
public class InsertionModelList implements Iterable<InsertionModel>
{
  private final static String COMBO_PARAM_ID = "ComboBoxWert";

  /**
   * Die Liste der {@link InsertionModel}s.
   */
  private List<InsertionModel> models = new LinkedList<InsertionModel>();

  /**
   * Liste aller {@link ItemListener}, die über Änderungen des Listeninhalts
   * informiert werden wollen.
   */
  private List<ItemListener> listeners = new ArrayList<ItemListener>(1);

  /**
   * Der FormularMax4000 zu dem diese InsertionModelList gehört.
   */
  private FormularMax4kController formularMax4000;

  /**
   * Erzeugt eine neue InsertionModelList.
   * 
   * @param formularMax4000
   *          der FormularMax4000 zu dem diese Liste gehört.
   * @author Matthias Benkmann (D-III-ITD 5.1)
   */
  public InsertionModelList(FormularMax4kController formularMax4000)
  {
    formularMax4000.addBroadcastListener(new MyBroadcastListener());
    this.formularMax4000 = formularMax4000;
  }

  /**
   * Fügt model dieser Liste hinzu.
   * 
   * @author Matthias Benkmann (D-III-ITD 5.1)
   */
  public void add(InsertionModel model)
  {
    int idx = models.size();
    models.add(idx, model);
    notifyListeners(model, idx, false);
  }

  /**
   * Löscht alle bestehenden InsertionModels aus der Liste.
   * 
   * @author Matthias Benkmann (D-III-ITD 5.1)
   */
  public void clear()
  {
    while (!models.isEmpty())
    {
      int index = models.size() - 1;
      InsertionModel model = models.remove(index);
      model.hasBeenRemoved();
      notifyListeners(model, index, true);
    }
  }

  /**
   * Liefert einen Iterator über alle {@link InsertionModel}s in dieser Liste.
   * ACHTUNG! Es dürfen keine Veränderungen über den Iterator (z.B.
   * {@link Iterator#remove()}) vorgenommen werden. Auch dürfen während der
   * Iteration keine Veränderungen an der InsertionModelList vorkommen, da der
   * Iterator direkt auf der internen Datenstruktur arbeitet und es daher zur
   * {@link java.util.ConcurrentModificationException} kommen würde.
   * 
   * @author Matthias Benkmann (D-III-ITD 5.1)
   */
  @Override
  public Iterator<InsertionModel> iterator()
  {
    return models.iterator();
  }

  /**
   * Bittet die InsertionModelList darum, das Element model aus sich zu entfernen
   * (falls es in der Liste ist).
   * 
   * @author Matthias Benkmann (D-III-ITD 5.1) TESTED
   */
  public void remove(InsertionModel model)
  {
    int index = models.indexOf(model);
    if (index < 0) return;
    models.remove(index);
    model.hasBeenRemoved();
    notifyListeners(model, index, true);
  }

  /**
   * Wird aufgerufen, nachdem mehrere Checkboxen in der Formularbeschreibung zu einer
   * ComboBox zusammengefasst wurden, damit die {@link InsertionModel}s, die vorher
   * auf die Checkboxen verwiesen haben entsprechend angepasst werden, so dass sie
   * jetzt auf die neue ComboBox verweisen und eine passende TRAFO bekommen.
   * 
   * @param desc
   * @author Matthias Benkmann (D-III-ITD 5.1) TESTED
   */
  public void mergeCheckboxesIntoCombobox(ComboboxMergeDescriptor desc)
  {
    FormControlModel combo = desc.combo;
    IDManager.ID comboIdd = combo.getId();
    if (comboIdd == null)
    {
      Logger.error(L.m("Programmfehler: Durch Merge erstellte ComboBox hat keine ID bekommen"));
      return;
    }
    String comboId = comboIdd.toString();
    Iterator<InsertionModel> iter = iterator();
    while (iter.hasNext())
    {
      try
      {
        InsertionModel4InsertXValue model =
          (InsertionModel4InsertXValue) iter.next();
        String comboValue = desc.mapCheckboxId2ComboboxEntry.get(model.getDataID());
        if (comboValue != null)
        {
          try
          {
            model.setDataID(comboId);
          }
          catch (UnknownIDException e)
          {
            Logger.error(L.m("Programmfehler"), e);
            return;
          }

          setMatchTrafo(model, comboValue);
        }
      }
      catch (ClassCastException x)
      {
        // skip this InsertionModel since it can't refer to a checkbox
      }
    }
  }

  /**
   * Setzt die TRAFO von model auf MATCH(VALUE "COMBO_PARAM_ID",
   * "re_escape(comboValue)").
   * 
   * @author Matthias Benkmann (D-III-ITD 5.1)
   */
  private void setMatchTrafo(InsertionModel model, String comboValue)
  {
    ConfigThingy trafoConf = new ConfigThingy("TRAFO");
    ConfigThingy matchConf = trafoConf.add("MATCH");
    matchConf.add("VALUE").add(COMBO_PARAM_ID);
    matchConf.add(re_escape(comboValue));
    FunctionSelectionAccess trafo = model.getTrafoAccess();
    trafo.setExpertFunction(trafoConf);
  }

  /**
   * Versucht TRAFOs von Einfügungen der ComboBox combo zu reparieren, die durch eine
   * Änderung der Werte-Liste von combo zerbrochen sind.
   * 
   * @author Matthias Benkmann (D-III-ITD 5.1) TESTED
   */
  public void fixComboboxInsertions(FormControlModel combo)
  {
    IDManager.ID comboId = combo.getId();
    if (comboId == null) return;
    Collection<String> items = combo.getItems();
    Collection<String> unusedItems = new HashSet<String>(items);
    Collection<InsertionModel> brokenInsertionModels = new Vector<InsertionModel>();
    Iterator<InsertionModel> iter = iterator();
    while (iter.hasNext())
    {
      InsertionModel4InsertXValue model;
      try
      {
        model = (InsertionModel4InsertXValue) iter.next();
      }
      catch (ClassCastException x)
      {
        model = null;
      }

      if (model != null && model.getDataID().equals(comboId))
      {
        FunctionSelectionAccess trafo = model.getTrafoAccess();
        if (!trafo.isExpert()) continue;
        ConfigThingy trafoConf = trafo.getExpertFunction();
        String regex;
        try
        {
          if (trafoConf.count() != 1
            || !trafoConf.getFirstChild().getName().equals("MATCH")) continue;

          trafoConf = trafoConf.getFirstChild();
          if (trafoConf.count() != 2
            || !trafoConf.getFirstChild().getName().equals("VALUE")) continue;

          trafoConf = trafoConf.getLastChild();
          if (trafoConf.count() != 0) continue;
          regex = trafoConf.toString();
        }
        catch (NodeNotFoundException e)
        {
          Logger.error(L.m("Kann nicht passieren"), e);
          return;
        }

        Pattern p = Pattern.compile(regex);
        boolean found = false;
        Iterator<String> itemsIter = items.iterator();
        while (itemsIter.hasNext())
        {
          String item = itemsIter.next();
          if (p.matcher(item).matches())
          {
            unusedItems.remove(item);
            found = true;
            break;
          }
        }
        if (!found) brokenInsertionModels.add(model);
      }
    }

    /*
     * Wenn wir ein unbenutztes Item haben, dann ändern wir alle TRAFOs, die broken
     * sind so, dass sie auf dieses matchen.
     */
    if (unusedItems.size() > 0)
    {
      String item = unusedItems.iterator().next();
      iter = brokenInsertionModels.iterator();
      while (iter.hasNext())
      {
        InsertionModel model = iter.next();
        setMatchTrafo(model, item);
      }
    }
  }

  /**
   * Liefert einen regulären Ausdruck, der genau den String str matcht (aber ohne ^
   * und $).
   * 
   * @author Matthias Benkmann (D-III-ITD 5.1) TESTED
   */
  private String re_escape(String str)
  {
    StringBuilder buffy = new StringBuilder();
    for (int i = 0; i < str.length(); ++i)
    {
      char ch = str.charAt(i);
      if (ch == ' ' || Character.isLetterOrDigit(ch))
        buffy.append(ch);
      else
      {
        buffy.append("\\u");
        String hexstr = Integer.toHexString(ch);
        for (int j = 4; j > hexstr.length(); --j)
          buffy.append('0');
        buffy.append(hexstr);
      }
    }
    return buffy.toString();
  }

  /**
   * Lässt alle in dieser Liste gespeicherten {@link InsertionModel}s ihre
   * zugehörigen Bookmarks updaten. Falls beim Update eines Bookmarks etwas
   * schiefgeht wird das entsprechende {@link InsertionModel} aus der Liste gelöscht.
   * Das Ausführen dieser Funktion triggert also potentiell einige Listener.
   * 
   * @param mapFunctionNameToConfigThingy
   *          bildet einen Funktionsnamen auf ein ConfigThingy ab, dessen Wurzel der
   *          Funktionsname ist und dessen Inhalt eine Funktionsdefinition. Wenn eine
   *          Einfügung mit einer TRAFO versehen ist, wird für das Aktualisieren des
   *          Bookmarks ein Funktionsname generiert, der noch nicht in dieser Map
   *          vorkommt und ein Mapping für diese Funktion wird in die Map eingefügt.
   *          Nach dem Aufruf von updateDocument() sind zu dieser Map also Einträge
   *          hinzugekommen für alle TRAFOs, die in den Einfügungen vorkommen.
   * @author Matthias Benkmann (D-III-ITD 5.1) TESTED
   */
  public void updateDocument(Map<String, ConfigThingy> mapFunctionNameToConfigThingy)
  {
    List<InsertionModel> defunct = new Vector<InsertionModel>();
    Iterator<InsertionModel> iter = models.iterator();
    while (iter.hasNext())
    {
      InsertionModel model = iter.next();
      if (!model.updateDocument(mapFunctionNameToConfigThingy)) defunct.add(model);
    }

    iter = defunct.iterator();
    while (iter.hasNext())
    {
      remove(iter.next());
    }
  }

  /**
   * listener wird über Änderungen der Liste informiert.
   * 
   * @author Matthias Benkmann (D-III-ITD 5.1)
   */
  public void addListener(ItemListener listener)
  {
    if (!listeners.contains(listener)) listeners.add(listener);
  }

  /**
   * Benachrichtigt alle ItemListener über das Hinzufügen oder Entfernen von model
   * zur bzw. aus der Liste an/von Index index.
   * 
   * @param removed
   *          falls true, wurde model entfernt, ansonsten hinzugefügt.
   * @author Matthias Benkmann (D-III-ITD 5.1)
   */
  private void notifyListeners(InsertionModel model, int index, boolean removed)
  {
    Iterator<ItemListener> iter = listeners.iterator();
    while (iter.hasNext())
    {
      ItemListener listener = iter.next();
      if (removed)
        listener.itemRemoved(model, index);
      else
        listener.itemAdded(model, index);
    }
  }

  /**
   * Interface für Klassen, die interessiert sind, zu erfahren, wenn sich die Liste
   * ändert.
   * 
   * @author Matthias Benkmann (D-III-ITD 5.1)
   */
  public static interface ItemListener
  {
    /**
     * Wird aufgerufen nachdem model zur Liste hinzugefügt wurde (an Index index).
     * 
     * @author Matthias Benkmann (D-III-ITD 5.1)
     */
    public void itemAdded(InsertionModel model, int index);

    /**
     * Wird aufgerufen, nachdem model aus der Liste entfernt wurde.
     * 
     * @param index
     *          der alte Index von model in der Liste.
     * @author Matthias Benkmann (D-III-ITD 5.1)
     */
    public void itemRemoved(InsertionModel model, int index);
  }

  private class MyBroadcastListener extends BroadcastListener
  {
    boolean insertionViewsSelected = false;

    @Override
    public void broadcastAllInsertionsViewSelected()
    {
      insertionViewsSelected = true;
    }

    @Override
    public void broadcastAllFormControlsViewSelected()
    {
      insertionViewsSelected = false;
    }

    @Override
    public void broadcastBookmarkSelection(Set<String> bookmarkNames)
    { // TESTED
      if (!insertionViewsSelected) return;
      boolean clearSelection = true;
      Iterator<InsertionModel> iter = models.iterator();
      while (iter.hasNext())
      {
        InsertionModel model = iter.next();
        if (bookmarkNames.contains(model.getName()))
        {
          formularMax4000.broadcast(new BroadcastObjectSelection(model, 1,
            clearSelection)
          {
            @Override
            public void sendTo(BroadcastListener listener)
            {
              listener.broadcastInsertionModelSelection(this);
            }
          });
          clearSelection = false;
        }
      }
    }
  }
}
