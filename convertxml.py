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

    template_begin_last_tab_position = '\t\t\t\t\t'
    for group in soup.tab.contents:
        if type(group) == Tag:      
            
            group_xml = []

            group_xml.append(f'''\n<Group id="{group['id']}">''')
            group_xml.append(f'''\t<Label resid="{group['id']}.label"/>''')       
            group_xml.append( '''\t<Icon>\n\t\t<bt:Image size="16" resid="Icon.16x16"/>\n\t\t<bt:Image size="32" resid="Icon.32x32"/>\n\t\t<bt:Image size="80" resid="Icon.80x80"/>\n\t</Icon>''')
            
            for button in group.contents:
                if type(button) == Tag:      
                    group_xml.append(f'''\t<Control xsi:type="Button" id="{button['id']}">''')
                    group_xml.append(f'''\t\t<Label resid="{button['id']}.label"/>''')       
                    group_xml.append( '''\t\t<Supertip>''')
                    group_xml.append(f'''\t\t\t<Title resid="{button['id']}.Title"/>''')       
                    group_xml.append(f'''\t\t\t<Description resid="{button['id']}.Desc"/>''')       
                    group_xml.append( '''\t\t</Supertip>''')


            group_xml.append('</Group>')  
            group_xml = [template_begin_last_tab_position + group for group in group_xml]
            f.write('\n'.join(group_xml))

            res_xml['ShortStrings'].append(f'''<bt:String id="{group['id']}.label" DefaultValue="{group['label']}"/>''')
        
    f.write(new_manifest_middle)

    f.write('\n\t\t\t<Resources>')

    f.write('\n\t\t\t</Resources>')

    f.write(new_manifest_end)

# print(res_xml)

# desired_element = soup1.find("YourElementName")  # Replace 'YourElementName' with the name of the desired element
# if desired_element:
#     extracted_content = str(desired_element.prettify())



