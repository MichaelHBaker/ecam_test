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
                    if line.name == 'button':
                        group_xml.append(f'''  <Control xsi:type="Button" id="{line['id']}">''')
                        group_xml.append(f'''    <Label resid="{line['id']}.label"/>''')       
                        group_xml.append( '''    <Supertip>''')
                        group_xml.append(f'''      <Title resid="{line['id']}.Title"/>''')       
                        group_xml.append(f'''      <Description resid="{line['id']}.Desc"/>''')       
                        group_xml.append( '''    </Supertip>''')
                        group_xml.append( '''    <Icon>''')
                        group_xml.append(f'''      <bt:Image size="16" resid="{line['imageMso']}.16x16"/>''')
                        group_xml.append(f'''      <bt:Image size="32" resid="{line['imageMso']}.32x32"/>''')
                        group_xml.append(f'''      <bt:Image size="80" resid="{line['imageMso']}.80x80"/>''')
                        group_xml.append( '''    </Icon>''')
                        group_xml.append( '''    <Action xsi:type="ExecuteFunction">''')
                        group_xml.append( '''      <FunctionName>OnAction_ECAM</FunctionName>''')
                        group_xml.append( '''    </Action>''')
                        group_xml.append( '''  </Control>''')
                        res_xml['Images'].append(f'''<bt:Image id="{line['imageMso']}.16x16" DefaultValue="https://localhost:3000/assets/IconSelectArea_16.png"/>''')
                        res_xml['Images'].append(f'''<bt:Image id="{line['imageMso']}.32x32" DefaultValue="https://localhost:3000/assets/IconSelectArea_32.png"/>''')
                        res_xml['Images'].append(f'''<bt:Image id="{line['imageMso']}.80x80" DefaultValue="https://localhost:3000/assets/IconSelectArea_80.png"/>''')
                        # res_xml['Images'].append(f'''<bt:Image id="{line['imageMso']}.16x16" DefaultValue="https://localhost:3000/assets/{line['imageMso']}_16.png"/>''')
                        # res_xml['Images'].append(f'''<bt:Image id="{line['imageMso']}.32x32" DefaultValue="https://localhost:3000/assets/{line['imageMso']}_32.png"/>''')
                        # res_xml['Images'].append(f'''<bt:Image id="{line['imageMso']}.80x80" DefaultValue="https://localhost:3000/assets/{line['imageMso']}_80.png"/>''')

                        res_xml['ShortStrings'].append(f'''<bt:String id="{line['id']}.label" DefaultValue="{line['label']}"/>''')
                        if 'screentip' in line.attrs:
                            res_xml['ShortStrings'].append(f'''<bt:String id="{line['id']}.title" DefaultValue="{line['screentip']}"/>''')
                            res_xml['ShortStrings'].append(f'''<bt:String id="{line['id']}.desc" DefaultValue="Click to {line['screentip']}"/>''')
                        else:
                            res_xml['ShortStrings'].append(f'''<bt:String id="{line['id']}.title" DefaultValue="{line['label']}"/>''')
                            res_xml['ShortStrings'].append(f'''<bt:String id="{line['id']}.desc" DefaultValue="Click to {line['label']}"/>''')

                    elif line.name == 'menu':
                        group_xml.append(f'''  <Control xsi:type="Menu" id="{line['id']}">''')
                        group_xml.append(f'''    <Label resid="{line['id']}.label"/>''')       
                        group_xml.append( '''    <Supertip>''')
                        group_xml.append(f'''      <Title resid="{line['id']}.Title"/>''')       
                        group_xml.append(f'''      <Description resid="{line['id']}.Desc"/>''')       
                        group_xml.append( '''    </Supertip>''')
                        group_xml.append( '''    <Icon>''')
                        group_xml.append(f'''      <bt:Image size="16" resid="{line['imageMso']}.16x16"/>''')
                        group_xml.append(f'''      <bt:Image size="32" resid="{line['imageMso']}.32x32"/>''')
                        group_xml.append(f'''      <bt:Image size="80" resid="{line['imageMso']}.80x80"/>''')
                        group_xml.append( '''    </Icon>''')
                        group_xml.append( '''    <Items>''')

                        res_xml['Images'].append(f'''<bt:Image id="{line['imageMso']}.16x16" DefaultValue="https://localhost:3000/assets/IconSelectArea_16.png"/>''')
                        res_xml['Images'].append(f'''<bt:Image id="{line['imageMso']}.32x32" DefaultValue="https://localhost:3000/assets/IconSelectArea_32.png"/>''')
                        res_xml['Images'].append(f'''<bt:Image id="{line['imageMso']}.80x80" DefaultValue="https://localhost:3000/assets/IconSelectArea_80.png"/>''')
                        # res_xml['Images'].append(f'''<bt:Image id="{line['imageMso']}.16x16" DefaultValue="https://localhost:3000/assets/{line['imageMso']}_16.png"/>''')
                        # res_xml['Images'].append(f'''<bt:Image id="{line['imageMso']}.32x32" DefaultValue="https://localhost:3000/assets/{line['imageMso']}_32.png"/>''')
                        # res_xml['Images'].append(f'''<bt:Image id="{line['imageMso']}.80x80" DefaultValue="https://localhost:3000/assets/{line['imageMso']}_80.png"/>''')

                        res_xml['ShortStrings'].append(f'''<bt:String id="{line['id']}.label" DefaultValue="{line['label']}"/>''')
                        if 'screentip' in line.attrs:
                            res_xml['ShortStrings'].append(f'''<bt:String id="{line['id']}.title" DefaultValue="{line['screentip']}"/>''')
                            res_xml['ShortStrings'].append(f'''<bt:String id="{line['id']}.desc" DefaultValue="Click to {line['screentip']}"/>''')
                        else:
                            res_xml['ShortStrings'].append(f'''<bt:String id="{line['id']}.title" DefaultValue="{line['label']}"/>''')
                            res_xml['ShortStrings'].append(f'''<bt:String id="{line['id']}.desc" DefaultValue="Click to {line['label']}"/>''')

                        for item in line.contents:
                            if type(item) == Tag:
                                group_xml.append(f'''      <Item id="{item['id']}">''')
                                group_xml.append(f'''        <Label resid="{item['id']}.label"/>''')       
                                group_xml.append( '''        <Supertip>''')
                                group_xml.append(f'''          <Title resid="{item['id']}.Title"/>''')       
                                group_xml.append(f'''          <Description resid="{item['id']}.Desc"/>''')       
                                group_xml.append( '''        </Supertip>''')
                                if 'imageMso' in item.attrs:
                                    group_xml.append( '''        <Icon>''')
                                    group_xml.append(f'''          <bt:Image size="16" resid="{item['imageMso']}.16x16"/>''')
                                    group_xml.append(f'''          <bt:Image size="32" resid="{item['imageMso']}.32x32"/>''')
                                    group_xml.append(f'''          <bt:Image size="80" resid="{item['imageMso']}.80x80"/>''')
                                    group_xml.append( '''        </Icon>''')
                                group_xml.append( '''        <Action xsi:type="ExecuteFunction">''')
                                group_xml.append( '''          <FunctionName>OnAction_ECAM</FunctionName>''')
                                group_xml.append( '''        </Action>''')
                                group_xml.append( '''      </Item>''')

                                if 'imageMso' in item:
                                    res_xml['Images'].append(f'''<bt:Image id="{item['imageMso']}.16x16" DefaultValue="https://localhost:3000/assets/IconSelectArea_16.png"/>''')
                                    res_xml['Images'].append(f'''<bt:Image id="{item['imageMso']}.32x32" DefaultValue="https://localhost:3000/assets/IconSelectArea_32.png"/>''')
                                    res_xml['Images'].append(f'''<bt:Image id="{item['imageMso']}.80x80" DefaultValue="https://localhost:3000/assets/IconSelectArea_80.png"/>''')
                                    # res_xml['Images'].append(f'''<bt:Image id="{item['imageMso']}.16x16" DefaultValue="https://localhost:3000/assets/{item['imageMso']}_16.png"/>''')
                                    # res_xml['Images'].append(f'''<bt:Image id="{item['imageMso']}.32x32" DefaultValue="https://localhost:3000/assets/{item['imageMso']}_32.png"/>''')
                                    # res_xml['Images'].append(f'''<bt:Image id="{item['imageMso']}.80x80" DefaultValue="https://localhost:3000/assets/{item['imageMso']}_80.png"/>''')
                                # else  an icon is required point each of these to the icon used for group that displays blank

                                res_xml['ShortStrings'].append(f'''<bt:String id="{item['id']}.label" DefaultValue="{item['label']}"/>''')
                                res_xml['ShortStrings'].append(f'''<bt:String id="{item['id']}.title" DefaultValue="{item['label']}"/>''')
                                res_xml['LongStrings'].append(f'''<bt:String id="{item['id']}.desc" DefaultValue="Click to {item['label']}"/>''')

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



