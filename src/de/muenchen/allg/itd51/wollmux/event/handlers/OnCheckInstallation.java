package de.muenchen.allg.itd51.wollmux.event.handlers;

import java.io.File;
import java.net.URI;
import java.net.URISyntaxException;
import java.text.DateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import com.sun.star.container.NoSuchElementException;
import com.sun.star.util.XStringSubstitution;

import de.muenchen.allg.afid.UNO;
import de.muenchen.allg.itd51.wollmux.ModalDialogs;
import de.muenchen.allg.itd51.wollmux.WollMuxFehlerException;
import de.muenchen.allg.itd51.wollmux.core.util.L;
import de.muenchen.allg.itd51.wollmux.core.util.Logger;

public class OnCheckInstallation extends BasicEvent
{
  @Override
  public void doit() throws WollMuxFehlerException
  {
    // Standardwerte für den Warndialog:
    boolean showdialog = true;
    String title = L.m("Mehrfachinstallation des WollMux");
    String msg = L.m(
        "Es wurden eine systemweite und eine benutzerlokale Installation des WollMux\n(oder Überreste von einer unvollständigen Deinstallation) gefunden.\nDiese Konstellation kann obskure Fehler verursachen.\n\nEntfernen Sie eine der beiden Installationen.\n\nDie wollmux.log enthält nähere Informationen zu den betroffenen Pfaden.");
    String logMsg = msg;

    // Infos der Installationen einlesen.
    List<OnCheckInstallation.WollMuxInstallationDescriptor> wmInsts = getInstallations();

    // Variablen recentInstPath / recentInstLastModified / shared / local
    // bestimmen
    String recentInstPath = "";
    Date recentInstLastModified = null;
    boolean shared = false;
    boolean local = false;
    for (OnCheckInstallation.WollMuxInstallationDescriptor desc : wmInsts)
    {
      shared = shared || desc.isShared;
      local = local || !desc.isShared;
      if (recentInstLastModified == null || desc.date.compareTo(recentInstLastModified) > 0)
      {
        recentInstLastModified = desc.date;
        recentInstPath = desc.path;
      }
    }

    // Variable wrongInstList bestimmen:
    String otherInstsList = "";
    for (OnCheckInstallation.WollMuxInstallationDescriptor desc : wmInsts)
    {
      if (!desc.path.equals(recentInstPath))
        otherInstsList += "- " + desc.path + "\n";
    }

    // Im Fehlerfall Dialog und Fehlermeldung bringen.
    if (local && shared)
    {

      // Variablen in msg evaluieren:
      DateFormat f = DateFormat.getDateInstance();
      msg = msg.replaceAll("\\$\\{RECENT_INST_PATH\\}", recentInstPath);
      msg = msg.replaceAll("\\$\\{RECENT_INST_LAST_MODIFIED\\}", f.format(recentInstLastModified));
      msg = msg.replaceAll("\\$\\{OTHER_INSTS_LIST\\}", otherInstsList);

      logMsg += "\n" + L.m("Die juengste WollMux-Installation liegt unter:") + "\n- " + recentInstPath + "\n"
          + L.m("Ausserdem wurden folgende WollMux-Installationen gefunden:") + "\n" + otherInstsList;
      Logger.error(logMsg);

      if (showdialog)
        ModalDialogs.showInfoModal(title, msg, 0);
    }
  }

  private static class WollMuxInstallationDescriptor
  {
    public String path;

    public Date date;

    public boolean isShared;

    public WollMuxInstallationDescriptor(String path, Date date, boolean isShared)
    {
      this.path = path;
      this.date = date;
      this.isShared = isShared;
    }

    @Override
    public String toString()
    {
      return path + " -- " + date + " shared:" + isShared;
    }
  }

  /**
   * Liefert eine {@link List} mit den aktuell auf dem System vorhandenen
   * WollMux-Installationen.
   * 
   * @author Christoph Lutz, Matthias Benkmann (D-III-ITD-D101)
   */
  private List<OnCheckInstallation.WollMuxInstallationDescriptor> getInstallations()
  {
    List<OnCheckInstallation.WollMuxInstallationDescriptor> wmInstallations = new ArrayList<OnCheckInstallation.WollMuxInstallationDescriptor>();

    // Installationspfade der Pakete bestimmen:
    String myPath = null; // user-Pfad
    String oooPath = null; // shared-Pfad
    String oooPathNew = null; // shared-Pfad (OOo 3.x)

    try
    {
      XStringSubstitution xSS = UNO.XStringSubstitution(UNO.createUNOService("com.sun.star.util.PathSubstitution"));

      // Benutzerinstallationspfad LiMux =
      // /home/<Benutzer>/.openoffice.org2/user
      // Benutzerinstallationspfad Windows 2000 C:/Dokumente und
      // Einstellungen/<Benutzer>/Anwendungsdaten/OpenOffice.org2/user
      myPath = xSS.substituteVariables("$(user)/uno_packages/cache/uno_packages/", true);
      // Sharedinstallationspfad LiMux /opt/openoffice.org2.0/
      // Sharedinstallationspfad Windows C:/Programme/OpenOffice.org<version>
      oooPath = xSS.substituteVariables("$(inst)/share/uno_packages/cache/uno_packages/", true);
      try
      {
        oooPathNew = xSS.substituteVariables("$(brandbaseurl)/share/uno_packages/cache/uno_packages/", true);
      } catch (NoSuchElementException e)
      {
        // OOo 2.x does not have $(brandbaseurl)
      }

    } catch (java.lang.Exception e)
    {
      Logger.error(e);
      return wmInstallations;
    }

    if (myPath == null || oooPath == null)
    {
      Logger.error(L.m("Bestimmung der Installationspfade für das WollMux-Paket fehlgeschlagen."));
      return wmInstallations;
    }

    findWollMuxInstallations(wmInstallations, myPath, false);
    findWollMuxInstallations(wmInstallations, oooPath, true);
    if (oooPathNew != null)
      findWollMuxInstallations(wmInstallations, oooPathNew, true);

    return wmInstallations;
  }

  /**
   * Sucht im übergebenen Pfad path nach Verzeichnissen die WollMux.oxt
   * enthalten und fügt die Information zu wmInstallations hinzu.
   * 
   * @author Bettina Bauer, Christoph Lutz, Matthias Benkmann (D-III-ITD-D101)
   */
  private static void findWollMuxInstallations(List<OnCheckInstallation.WollMuxInstallationDescriptor> wmInstallations,
      String path, boolean isShared)
  {
    URI uriPath;
    uriPath = null;
    try
    {
      uriPath = new URI(path);
    } catch (URISyntaxException e)
    {
      Logger.error(e);
      return;
    }

    File[] installedPackages = new File(uriPath).listFiles();
    if (installedPackages != null)
    {
      // iterieren über die Installationsverzeichnisse mit automatisch
      // generierten Namen (z.B. 31GFBd_)
      for (int i = 0; i < installedPackages.length; i++)
      {
        if (installedPackages[i].isDirectory())
        {
          File dir = installedPackages[i];
          File[] dateien = dir.listFiles();
          for (int j = 0; j < dateien.length; j++)
          {
            // Wenn das Verzeichnis WollMux.oxt enthält, speichern des
            // Verzeichnisnames und des Verzeichnisdatum in einer HashMap
            if (dateien[j].isDirectory() && dateien[j].getName().startsWith("WollMux."))
            {
              // Name des Verzeichnis in dem sich WollMux.oxt befindet
              String directoryName = dateien[j].getAbsolutePath();
              // Datum des Verzeichnis in dem sich WollMux.oxt befindet
              Date directoryDate = new Date(dateien[j].lastModified());
              wmInstallations.add(new WollMuxInstallationDescriptor(directoryName, directoryDate, isShared));
            }
          }
        }
      }
    }
  }

  @Override
  public String toString()
  {
    return this.getClass().getSimpleName() + "()";
  }
}