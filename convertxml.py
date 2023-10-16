from bs4 import BeautifulSoup, Tag

with open("manifest_template_begin.xml", "r") as f:
    new_manifest_begin = f.read()
with open("manifest_template_middle.xml", "r") as f:
    new_manifest_middle = f.read()
with open("manifest_template_end.xml", "r") as f:
    new_manifest_end = f.read()

with open("ecam_vba_ribbon.xml", "r") as f:
    old_manifest = f.read()
    soup = BeautifulSoup(old_manifest, 'lxml-xml')
    res_xml = {'Images':[], 'Urls':[], 'ShortStrings':[], 'LongStrings':[]}

with open("manifest_bs4.xml", "w") as f:
    f.write(new_manifest_begin)

    template_begin_last_tab_position = '  ' * 6

    for group in soup.tab.contents:
        if type(group) == Tag:      
            
            group_xml = []

            group_xml.append(f'''<Group id="{group['id']}">''')
            group_xml.append(f'''  <Label resid="{group['id']}.label"/>''')       
            group_xml.append( '''  <Icon>''')
            group_xml.append( '''    <bt:Image size="16" resid="Icon.16x16"/>''')
            group_xml.append( '''    <bt:Image size="32" resid="Icon.32x32"/>''')
            group_xml.append( '''    <bt:Image size="80" resid="Icon.80x80"/>''')
            group_xml.append( '''  </Icon>''')
            
            for line in group.contents:
                if type(line) == Tag:
                    # button_tag
                    # menu_open_tag
                    # menu_close_tag
                    # menu_button_tag
                    if line.name == 'button':
                        group_xml.append(f'''  <Control xsi:type="Button" id="{line['id']}">''')
                        group_xml.append(f'''    <Label resid="{line['id']}.label"/>''')       
                        group_xml.append( '''    <Supertip>''')
                        group_xml.append(f'''      <Title resid="{line['id']}.Title"/>''')       
                        group_xml.append(f'''      <Description resid="{line['id']}.Desc"/>''')       
                        group_xml.append( '''    </Supertip>''')
                        group_xml.append( '''    <Icon>''')
                        group_xml.append(f'''      <bt:Image size="16" resid="IconSelectArea.16x16"/>''')
                        group_xml.append(f'''      <bt:Image size="32" resid="IconSelectArea.32x32"/>''')
                        group_xml.append(f'''      <bt:Image size="80" resid="IconSelectArea.80x80"/>''')
                        group_xml.append( '''    </Icon>''')
                        group_xml.append( '''    <Action>''')
                        group_xml.append( '''      <FunctionName>OnAction_ECAM</FunctionName>''')
                        group_xml.append( '''    </Action>''')
                        group_xml.append( '''  </Control>''')
                    elif line.name == 'menu':
                        group_xml.append(f'''  <Control xsi:type="Menu" id="{line['id']}">''')
                        group_xml.append(f'''    <Label resid="{line['id']}.label"/>''')       
                        group_xml.append( '''    <Supertip>''')
                        group_xml.append(f'''      <Title resid="{line['id']}.Title"/>''')       
                        group_xml.append(f'''      <Description resid="{line['id']}.Desc"/>''')       
                        group_xml.append( '''    </Supertip>''')
                        group_xml.append( '''    <Icon>''')
                        group_xml.append(f'''      <bt:Image size="16" resid="IconSelectArea.16x16"/>''')
                        group_xml.append(f'''      <bt:Image size="32" resid="IconSelectArea.32x32"/>''')
                        group_xml.append(f'''      <bt:Image size="80" resid="IconSelectArea.80x80"/>''')
                        group_xml.append( '''    </Icon>''')
                        group_xml.append( '''    </Items>''')
                        for item in line.contents:
                            if type(item) == Tag:
                                group_xml.append(f'''      <Item id="{item['id']}">''')
                                group_xml.append(f'''        <Label resid="{item['id']}.label"/>''')       
                                group_xml.append( '''        <Supertip>''')
                                group_xml.append(f'''          <Title resid="{item['id']}.Title"/>''')       
                                group_xml.append(f'''          <Description resid="{item['id']}.Desc"/>''')       
                                group_xml.append( '''        </Supertip>''')
                                group_xml.append( '''        <Icon>''')
                                group_xml.append(f'''          <bt:Image size="16" resid="IconSelectArea.16x16"/>''')
                                group_xml.append(f'''          <bt:Image size="32" resid="IconSelectArea.32x32"/>''')
                                group_xml.append(f'''          <bt:Image size="80" resid="IconSelectArea.80x80"/>''')
                                group_xml.append( '''        </Icon>''')
                                group_xml.append( '''        <Action>''')
                                group_xml.append( '''          <FunctionName>OnAction_ECAM</FunctionName>''')
                                group_xml.append( '''        </Action>''')
                                group_xml.append( '''      </Item>''')
                        group_xml.append( '''        </Items>''')
                        group_xml.append( '''  </Control>''')
                        



            group_xml.append('</Group>\n')  
            group_xml = [template_begin_last_tab_position + group for group in group_xml]
            f.write('\n'.join(group_xml))

            res_xml['ShortStrings'].append(f'''<bt:String id="{group['id']}.label" DefaultValue="{group['label']}"/>''')
        
    f.write(new_manifest_middle)

    f.write('\n    <Resources>')
    for res_type in res_xml:
        f.write(f'\n      <bt:{res_type}>')
        for res in res_xml[res_type]:
            f.write(f'\n        {res}')    
        f.write(f'\n      </bt:{res_type}>')

    f.write('\n    </Resources>\n')

    f.write(new_manifest_end)

# print(res_xml)

# desired_element = soup1.find("YourElementName")  # Replace 'YourElementName' with the name of the desired element
# if desired_element:
#     extracted_content = str(desired_element.prettify())



