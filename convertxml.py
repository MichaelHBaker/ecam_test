from bs4 import BeautifulSoup

with open("manifest_template.xml", "r") as f:
    new_manifest = f.read()

with open("ecam_vba_ribbon.xml", "r") as f:
    old_manifest = f.read()

soup = BeautifulSoup(old_manifest, 'lxml-xml')

for group in soup.tab.children:
    print(group)

# desired_element = soup1.find("YourElementName")  # Replace 'YourElementName' with the name of the desired element
# if desired_element:
#     extracted_content = str(desired_element.prettify())


# combined_content = extracted_content + "\n" + str(soup2)

# with open("manifest_bs4.xml", "w") as f:
#     f.write(combined_content)

