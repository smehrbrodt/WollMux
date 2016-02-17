package de.muenchen.allg.itd51.wollmux;

import com.sun.star.container.XSet;
import com.sun.star.frame.TerminationVetoException;
import com.sun.star.frame.XTerminateListener;
import com.sun.star.lang.EventObject;
import com.sun.star.ui.XUIElementFactoryRegistration;
import com.sun.star.uno.UnoRuntime;

import de.muenchen.allg.afid.UNO;
import de.muenchen.allg.itd51.wollmux.comp.WollMux;
import de.muenchen.allg.itd51.wollmux.sidebar.WollMuxSidebarFactory;

/**
 * Über diese Klasse kann der WollMux zum Debuggen in der lokalen JVM gestartet
 * werden, ohne dass die Extension in OpenOffice/LibreOffice installiert ist. Bisher
 * haben wir für diesen Zweck die WollMuxBar mit der Konfigurationsoption
 * ALLOW_EXTERNAL_WOLLMUX "true" verwendet. Damit das Debuggen auch ohne WollMuxBar
 * möglich ist, wurde diese Klasse eingefügt. Zusammen mit dem neuen ant build-target
 * WollMux.oxt-ButtonsOnly kann so das Debugging vereinfacht werden.
 * 
 * Verwendung: Diese Main-Methode einfach per Debugger starten.
 * 
 * @author Christoph
 * 
 */
public class DebugExternalWollMux
{
  public static void main(String[] args) throws Exception
  {
    // Logger zum Debuggen auf stdout initialisieren und die zukünftigen
    // Logger-Einstellungen aus der wollmuxconfig ignorieren.
    Logger.init(System.out, Logger.DEBUG);
    Logger.setIgnoreInit(true);

    UNO.init();

    UNO.desktop.addTerminateListener(new XTerminateListener()
    {

      @Override
      public void disposing(EventObject arg0)
      {}

      @Override
      public void queryTermination(EventObject arg0) throws TerminationVetoException
      {}

      @Override
      public void notifyTermination(EventObject arg0)
      {
        Logger.log(L.m("Desktop wurde geschlossen - beende DebugExternalWollMux"));
        System.exit(0);
      }
    });

    XUIElementFactoryRegistration factoryRegistration =
      UnoRuntime.queryInterface(
        XUIElementFactoryRegistration.class,
        UNO.defaultContext.getServiceManager().createInstanceWithContext(
          "com.sun.star.ui.UIElementFactoryManager", UNO.defaultContext));

    factoryRegistration.registerFactory("toolpanel", "WollMuxSidebarFactory", null,
      "de.muenchen.allg.itd51.wollmux.sidebar.WollMuxSidebarFactory");
    XSet set =
      UnoRuntime.queryInterface(XSet.class, UNO.defaultContext.getServiceManager());
    set.insert(WollMux.__getComponentFactory(WollMux.class.getName()));
    set.insert(WollMux.__getComponentFactory(WollMuxSidebarFactory.class.getName()));

    WollMux wm = new WollMux(UNO.defaultContext);
  }
}
