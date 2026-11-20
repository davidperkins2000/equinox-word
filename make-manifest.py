# Writes word/manifest.xml for the hosting address given (default: GitHub Pages for davidperkins2000/equinox-word)
import sys, uuid
BASE = (sys.argv[1] if len(sys.argv) > 1 else 'https://davidperkins2000.github.io/equinox-word').rstrip('/') + '/word'
ID = '6f1c2e4a-8d3b-4e5f-9a7c-2b1d0e3f4a51'   # fixed: Word knows the add-in by this id
VER = '1.0.4.0'
# The Equinox tab's buttons: each opens the pane with ?do=<action>, which the pane runs once it knows the matter
B = {  # resid: (label, tip, icon file)
  'Pane':   ('Equinox', 'Show the Equinox pane for this document.', 'icon'),
  'File':   ('File to Equinox', 'File this document to the matter in Equinox.', 'file'),
  'Letter': ('New letter', 'Open a letter template as a new document for this matter.', 'letter'),
  'Fill':   ('Fill fields', 'Fill the #CODES# in this document from Equinox.', 'fill'),
  'Pdf':    ('Save PDF', "Save this document as a PDF in the matter's Drive folder.", 'pdf'),
  'Email':  ('Email PDF', 'Save the PDF and make a Gmail draft with it attached, to be filed to Equinox once sent.', 'email'),
}
ICONS = {'File': 'file', 'Letter': 'letter', 'Fill': 'fill', 'Pdf': 'pdf', 'Email': 'email'}
def button(cid, key, icon, do):
    ic = 'Icon' if icon == 'Icon' else icon
    url = 'Pane.Url' if not do else 'Pane.' + key + '.Url'
    return f'''                <Control xsi:type="Button" id="{cid}">
                  <Label resid="{key}.Btn"/>
                  <Supertip><Title resid="{key}.Btn"/><Description resid="{key}.Tip"/></Supertip>
                  <Icon><bt:Image size="16" resid="{ic}.16"/><bt:Image size="32" resid="{ic}.32"/><bt:Image size="80" resid="{ic}.80"/></Icon>
                  <Action xsi:type="ShowTaskpane">
                    <TaskpaneId>Office.AutoShowTaskpaneWithDocument</TaskpaneId>
                    <SourceLocation resid="{url}"/>
                  </Action>
                </Control>'''
def images():
    return '\n'.join(f'        <bt:Image id="{k}.{s}" DefaultValue="{BASE}/assets/{v}-{s}.png"/>' for k, v in ICONS.items() for s in (16, 32, 80))
def urls():
    return '\n'.join(f'        <bt:Url id="Pane.{k}.Url" DefaultValue="{BASE}/taskpane.html?do={k.lower()}"/>' for k in ICONS)
def labels():
    return '\n'.join(f'        <bt:String id="{k}.Btn" DefaultValue="{v[0]}"/>' for k, v in B.items())
def tips():
    return '\n'.join(f'        <bt:String id="{k}.Tip" DefaultValue="{v[1]}"/>' for k, v in B.items())

m = f'''<?xml version="1.0" encoding="UTF-8"?>
<!-- Equinox for Word (Eureka IP), feature 13.7. Sideload: copy this file to
     ~/Library/Containers/com.microsoft.Word/Data/Documents/wef/ and restart Word. -->
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
           xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
           xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0"
           xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides"
           xsi:type="TaskPaneApp">
  <Id>{ID}</Id>
  <Version>{VER}</Version>
  <ProviderName>Eureka IP</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="Equinox"/>
  <Description DefaultValue="Equinox matter details, filing and check-out for Word documents in the matters' Drive folders."/>
  <IconUrl DefaultValue="{BASE}/assets/icon-32.png"/>
  <HighResolutionIconUrl DefaultValue="{BASE}/assets/icon-64.png"/>
  <SupportUrl DefaultValue="https://eurekaip.com.au"/>
  <AppDomains>
    <AppDomain>https://script.google.com</AppDomain>
    <AppDomain>https://script.googleusercontent.com</AppDomain>
  </AppDomains>
  <Hosts><Host Name="Document"/></Hosts>
  <Requirements><Sets DefaultMinVersion="1.1"><Set Name="WordApi" MinVersion="1.1"/></Sets></Requirements>
  <DefaultSettings><SourceLocation DefaultValue="{BASE}/taskpane.html"/></DefaultSettings>
  <Permissions>ReadWriteDocument</Permissions>
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Hosts>
      <Host xsi:type="Document">
        <DesktopFormFactor>
          <GetStarted>
            <Title resid="GetStarted.Title"/>
            <Description resid="GetStarted.Description"/>
            <LearnMoreUrl resid="Home.Url"/>
          </GetStarted>
          <FunctionFile resid="Commands.Url"/>
          <ExtensionPoint xsi:type="PrimaryCommandSurface">
            <OfficeTab id="TabHome">
              <Group id="Eqx.Group">
                <Label resid="Group.Label"/>
                <Icon><bt:Image size="16" resid="Icon.16"/><bt:Image size="32" resid="Icon.32"/><bt:Image size="80" resid="Icon.80"/></Icon>
                <Control xsi:type="Button" id="Eqx.Open">
                  <Label resid="Button.Label"/>
                  <Supertip><Title resid="Button.Label"/><Description resid="Button.Tip"/></Supertip>
                  <Icon><bt:Image size="16" resid="Icon.16"/><bt:Image size="32" resid="Icon.32"/><bt:Image size="80" resid="Icon.80"/></Icon>
                  <Action xsi:type="ShowTaskpane">
                    <TaskpaneId>Office.AutoShowTaskpaneWithDocument</TaskpaneId>
                    <SourceLocation resid="Pane.Url"/>
                  </Action>
                </Control>
              </Group>
            </OfficeTab>
            <CustomTab id="Eqx.Tab">
              <Group id="Eqx.MatterGroup">
                <Label resid="Matter.Label"/>
                <Icon><bt:Image size="16" resid="Icon.16"/><bt:Image size="32" resid="Icon.32"/><bt:Image size="80" resid="Icon.80"/></Icon>
{button('Eqx.TabOpen', 'Pane', 'Icon', '')}
{button('Eqx.File', 'File', 'File', 'file')}
              </Group>
              <Group id="Eqx.LetterGroup">
                <Label resid="Letter.Label"/>
                <Icon><bt:Image size="16" resid="Letter.16"/><bt:Image size="32" resid="Letter.32"/><bt:Image size="80" resid="Letter.80"/></Icon>
{button('Eqx.Letter', 'Letter', 'Letter', 'letter')}
{button('Eqx.Fill', 'Fill', 'Fill', 'fill')}
{button('Eqx.Pdf', 'Pdf', 'Pdf', 'pdf')}
{button('Eqx.Email', 'Email', 'Email', 'email')}
              </Group>
              <Label resid="Tab.Label"/>
            </CustomTab>
          </ExtensionPoint>
        </DesktopFormFactor>
      </Host>
    </Hosts>
    <Resources>
      <bt:Images>
        <bt:Image id="Icon.16" DefaultValue="{BASE}/assets/icon-16.png"/>
        <bt:Image id="Icon.32" DefaultValue="{BASE}/assets/icon-32.png"/>
        <bt:Image id="Icon.80" DefaultValue="{BASE}/assets/icon-80.png"/>
{images()}
      </bt:Images>
      <bt:Urls>
        <bt:Url id="Pane.Url" DefaultValue="{BASE}/taskpane.html"/>
        <bt:Url id="Commands.Url" DefaultValue="{BASE}/commands.html"/>
{urls()}
        <bt:Url id="Home.Url" DefaultValue="https://eurekaip.com.au"/>
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="GetStarted.Title" DefaultValue="Equinox is ready"/>
        <bt:String id="Group.Label" DefaultValue="Equinox"/>
        <bt:String id="Button.Label" DefaultValue="Equinox"/>
        <bt:String id="Tab.Label" DefaultValue="Equinox"/>
        <bt:String id="Matter.Label" DefaultValue="Matter"/>
        <bt:String id="Letter.Label" DefaultValue="Letter"/>
{labels()}
      </bt:ShortStrings>
      <bt:LongStrings>
        <bt:String id="GetStarted.Description" DefaultValue="Open Home, then Equinox, to see the document's matter."/>
        <bt:String id="Button.Tip" DefaultValue="Show the Equinox matter for this document: details, File to Equinox, Open from Equinox and Insert summary."/>
{tips()}
      </bt:LongStrings>
    </Resources>
  </VersionOverrides>
</OfficeApp>
'''
open('word/manifest.xml', 'w').write(m)
print('manifest for', BASE)
