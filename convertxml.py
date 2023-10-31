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
 
    # CustomTab Label is set on first line of manifest_template_middle.xml
    res_xml['ShortStrings'].append( '''<bt:String id="tbECAM_JS.L" DefaultValue="ECAM_JS" />''')
    
    res_xml['Urls'].append( '''<bt:Url id="Commands.Url" DefaultValue="https://localhost:3000/commands.html"/>''')

with open("manifest_bs4.xml", "w") as f:
    f.write(new_manifest_begin)

    template_begin_last_tab_position = '  ' * 7

    for group in soup.tab.contents:
        if type(group) == Tag:      
            
            group_xml = []
            if len(group['id']) > 30:
                group['id'] = group['id'][:30]
            group_xml.append(f'''<Group id="{group['id']}">''')
            group_xml.append(f'''  <Label resid="{group['id']}.L"/>''')       
            group_xml.append( '''  <Icon>''')
            group_xml.append(f'''    <bt:Image size="16" resid="{group['id']}.1"/>''')
            group_xml.append(f'''    <bt:Image size="32" resid="{group['id']}.3"/>''')
            group_xml.append(f'''    <bt:Image size="80" resid="{group['id']}.8"/>''')
            group_xml.append( '''  </Icon>''')
            res_xml['Images'].append(f'''<bt:Image id="{group['id']}.1" DefaultValue="https://localhost:3000/assets/IconSelectArea_16.png"/>''')
            res_xml['Images'].append(f'''<bt:Image id="{group['id']}.3" DefaultValue="https://localhost:3000/assets/IconSelectArea_32.png"/>''')
            res_xml['Images'].append(f'''<bt:Image id="{group['id']}.8" DefaultValue="https://localhost:3000/assets/IconSelectArea_80.png"/>''')
            # res_xml['Images'].append(f'''<bt:Image id="{group['id']}.1" DefaultValue="https://localhost:3000/assets/{group['id']}_16.png"/>''')
            # res_xml['Images'].append(f'''<bt:Image id="{group['id']}.3" DefaultValue="https://localhost:3000/assets/{group['id']}_32.png"/>''')
            # res_xml['Images'].append(f'''<bt:Image id="{group['id']}.8" DefaultValue="https://localhost:3000/assets/{group['id']}_80.png"/>''')
            res_xml['ShortStrings'].append(f'''<bt:String id="{group['id']}.L" DefaultValue="{group['label']}"/>''')
            
            for line in group.contents:
                if type(line) == Tag:
                    if line.name == 'button':
                        if len(line['id']) > 30:
                            line['id'] = line['id'][:30]
                        if 'imageMso' in line.attrs and len(line['imageMso']) > 30:
                            line['imageMso'] = line['imageMso'][:30]

                        group_xml.append(f'''  <Control xsi:type="Button" id="{line['id']}">''')
                        group_xml.append(f'''    <Label resid="{line['id']}.L"/>''')       
                        group_xml.append( '''    <Supertip>''')
                        group_xml.append(f'''      <Title resid="{line['id']}.T"/>''')       
                        group_xml.append(f'''      <Description resid="{line['id']}.D"/>''')       
                        group_xml.append( '''    </Supertip>''')
                        group_xml.append( '''    <Icon>''')
                        group_xml.append(f'''      <bt:Image size="16" resid="{line['imageMso']}.1"/>''')
                        group_xml.append(f'''      <bt:Image size="32" resid="{line['imageMso']}.3"/>''')
                        group_xml.append(f'''      <bt:Image size="80" resid="{line['imageMso']}.8"/>''')
                        group_xml.append( '''    </Icon>''')
                        group_xml.append( '''    <Action xsi:type="ExecuteFunction">''')
                        group_xml.append( '''      <FunctionName>OnAction_ECAM</FunctionName>''')
                        group_xml.append( '''    </Action>''')
                        group_xml.append( '''  </Control>''')
                        res_xml['Images'].append(f'''<bt:Image id="{line['imageMso']}.1" DefaultValue="https://localhost:3000/assets/IconSelectArea_16.png"/>''')
                        res_xml['Images'].append(f'''<bt:Image id="{line['imageMso']}.3" DefaultValue="https://localhost:3000/assets/IconSelectArea_32.png"/>''')
                        res_xml['Images'].append(f'''<bt:Image id="{line['imageMso']}.8" DefaultValue="https://localhost:3000/assets/IconSelectArea_80.png"/>''')
                        # res_xml['Images'].append(f'''<bt:Image id="{line['imageMso']}.1" DefaultValue="https://localhost:3000/assets/{line['imageMso']}_16.png"/>''')
                        # res_xml['Images'].append(f'''<bt:Image id="{line['imageMso']}.3" DefaultValue="https://localhost:3000/assets/{line['imageMso']}_32.png"/>''')
                        # res_xml['Images'].append(f'''<bt:Image id="{line['imageMso']}.8" DefaultValue="https://localhost:3000/assets/{line['imageMso']}_80.png"/>''')

                        res_xml['ShortStrings'].append(f'''<bt:String id="{line['id']}.L" DefaultValue="{line['label']}"/>''')
                        if 'screentip' in line.attrs:
                            res_xml['ShortStrings'].append(f'''<bt:String id="{line['id']}.T" DefaultValue="{line['screentip']}"/>''')
                            res_xml['LongStrings'].append(f'''<bt:String id="{line['id']}.D" DefaultValue="Click to {line['screentip']}"/>''')
                        else:
                            res_xml['ShortStrings'].append(f'''<bt:String id="{line['id']}.T" DefaultValue="{line['label']}"/>''')
                            res_xml['LongStrings'].append(f'''<bt:String id="{line['id']}.D" DefaultValue="Click to {line['label']}"/>''')

                    elif line.name == 'menu':
                        if len(line['id']) > 30:
                            line['id'] = line['id'][:30]
                        if 'imageMso' in line.attrs and len(line['imageMso']) > 30:
                            line['imageMso'] = line['imageMso'][:30]

                        group_xml.append(f'''  <Control xsi:type="Menu" id="{line['id']}">''')
                        group_xml.append(f'''    <Label resid="{line['id']}.L"/>''')       
                        group_xml.append( '''    <Supertip>''')
                        group_xml.append(f'''      <Title resid="{line['id']}.T"/>''')       
                        group_xml.append(f'''      <Description resid="{line['id']}.D"/>''')       
                        group_xml.append( '''    </Supertip>''')
                        group_xml.append( '''    <Icon>''')
                        group_xml.append(f'''      <bt:Image size="16" resid="{line['imageMso']}.1"/>''')
                        group_xml.append(f'''      <bt:Image size="32" resid="{line['imageMso']}.3"/>''')
                        group_xml.append(f'''      <bt:Image size="80" resid="{line['imageMso']}.8"/>''')
                        group_xml.append( '''    </Icon>''')
                        group_xml.append( '''    <Items>''')

                        res_xml['Images'].append(f'''<bt:Image id="{line['imageMso']}.1" DefaultValue="https://localhost:3000/assets/IconSelectArea_16.png"/>''')
                        res_xml['Images'].append(f'''<bt:Image id="{line['imageMso']}.3" DefaultValue="https://localhost:3000/assets/IconSelectArea_32.png"/>''')
                        res_xml['Images'].append(f'''<bt:Image id="{line['imageMso']}.8" DefaultValue="https://localhost:3000/assets/IconSelectArea_80.png"/>''')
                        # res_xml['Images'].append(f'''<bt:Image id="{line['imageMso']}.1" DefaultValue="https://localhost:3000/assets/{line['imageMso']}_16.png"/>''')
                        # res_xml['Images'].append(f'''<bt:Image id="{line['imageMso']}.3" DefaultValue="https://localhost:3000/assets/{line['imageMso']}_32.png"/>''')
                        # res_xml['Images'].append(f'''<bt:Image id="{line['imageMso']}.8" DefaultValue="https://localhost:3000/assets/{line['imageMso']}_80.png"/>''')

                        res_xml['ShortStrings'].append(f'''<bt:String id="{line['id']}.L" DefaultValue="{line['label']}"/>''')
                        if 'screentip' in line.attrs:
                            res_xml['ShortStrings'].append(f'''<bt:String id="{line['id']}.T" DefaultValue="{line['screentip']}"/>''')
                            res_xml['LongStrings'].append(f'''<bt:String id="{line['id']}.D" DefaultValue="Click to {line['screentip']}"/>''')
                        else:
                            res_xml['ShortStrings'].append(f'''<bt:String id="{line['id']}.T" DefaultValue="{line['label']}"/>''')
                            res_xml['LongStrings'].append(f'''<bt:String id="{line['id']}.D" DefaultValue="Click to {line['label']}"/>''')

                        for item in line.contents:
                            if type(item) == Tag:

                                if item.name == 'button':
                                    if len(item['id']) > 30:
                                        item['id'] = item['id'][:30]
                                    if 'imageMso' in item.attrs and len(item['imageMso']) > 30:
                                        item['imageMso'] = item['imageMso'][:30]
                                group_xml.append(f'''      <Item id="{item['id']}">''')
                                group_xml.append(f'''        <Label resid="{item['id']}.L"/>''')       
                                group_xml.append( '''        <Supertip>''')
                                group_xml.append(f'''          <Title resid="{item['id']}.T"/>''')       
                                group_xml.append(f'''          <Description resid="{item['id']}.D"/>''')       
                                group_xml.append( '''        </Supertip>''')
                                if 'imageMso' in item.attrs:
                                    group_xml.append( '''        <Icon>''')
                                    group_xml.append(f'''          <bt:Image size="16" resid="{item['imageMso']}.1"/>''')
                                    group_xml.append(f'''          <bt:Image size="32" resid="{item['imageMso']}.3"/>''')
                                    group_xml.append(f'''          <bt:Image size="80" resid="{item['imageMso']}.8"/>''')
                                    group_xml.append( '''        </Icon>''')
                                group_xml.append( '''        <Action xsi:type="ExecuteFunction">''')
                                group_xml.append( '''          <FunctionName>OnAction_ECAM</FunctionName>''')
                                group_xml.append( '''        </Action>''')
                                group_xml.append( '''      </Item>''')

                                if 'imageMso' in item.attrs:
                                    res_xml['Images'].append(f'''<bt:Image id="{item['imageMso']}.1" DefaultValue="https://localhost:3000/assets/IconSelectArea_16.png"/>''')
                                    res_xml['Images'].append(f'''<bt:Image id="{item['imageMso']}.3" DefaultValue="https://localhost:3000/assets/IconSelectArea_32.png"/>''')
                                    res_xml['Images'].append(f'''<bt:Image id="{item['imageMso']}.8" DefaultValue="https://localhost:3000/assets/IconSelectArea_80.png"/>''')
                                    # res_xml['Images'].append(f'''<bt:Image id="{item['imageMso']}.1" DefaultValue="https://localhost:3000/assets/{item['imageMso']}_16.png"/>''')
                                    # res_xml['Images'].append(f'''<bt:Image id="{item['imageMso']}.3" DefaultValue="https://localhost:3000/assets/{item['imageMso']}_32.png"/>''')
                                    # res_xml['Images'].append(f'''<bt:Image id="{item['imageMso']}.8" DefaultValue="https://localhost:3000/assets/{item['imageMso']}_80.png"/>''')
                                # else  an icon is required point each of these to the icon used for group that displays blank

                                res_xml['ShortStrings'].append(f'''<bt:String id="{item['id']}.L" DefaultValue="{item['label']}"/>''')
                                res_xml['ShortStrings'].append(f'''<bt:String id="{item['id']}.T" DefaultValue="{item['label']}"/>''')
                                res_xml['LongStrings'].append(f'''<bt:String id="{item['id']}.D" DefaultValue="Click to {item['label']}"/>''')

                        group_xml.append( '''    </Items>''')
                        group_xml.append( '''  </Control>''')
                        



            group_xml.append('</Group>\n')  
            group_xml = [template_begin_last_tab_position + group for group in group_xml]
            f.write('\n'.join(group_xml))

        
    f.write(new_manifest_middle)

    f.write('\n    <Resources>')
    for res_type in res_xml:
        if res_xml[res_type] != []:
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



