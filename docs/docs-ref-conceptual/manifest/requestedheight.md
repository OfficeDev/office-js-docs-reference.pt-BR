# <a name="requestedheight-element"></a>Elemento RequestedHeight

Especifica a altura inicial (em pixels) de um conteúdo suplemento ou o suplemento de email. 

**Tipo de suplemento:** Conteúdo de email

## <a name="syntax"></a>Sintaxe

```XML
<RequestedHeight>integer</RequestedHeight>
```

## <a name="contained-in"></a>Contidos em

- [DefaultSettings](defaultsettings.md) (Complementos de conteúdo) com um valor que pode ser entre 32 e 1000
- [DesktopSettings](desktopsettings.md) e [TabletSettings](tabletsettings.md) (Mail suplementos) com um valor que pode ser entre 32 e 450
- [ExtensionPoint](extensionpoint.md) (Mail contextual suplementos) com um valor que pode ser entre 140 e 450 para o ponto de extensão **DetectedEntity** e entre 32 e 450 para o ponto de extensão **CustomPane**