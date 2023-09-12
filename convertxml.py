from bs4 import BeautifulSoup

# Step 1: Extract lines from the first XML file

with open("first_xml_file.xml", "r") as f:
    contents = f.read()

soup1 = BeautifulSoup(contents, 'lxml-xml')

desired_element = soup1.find("YourElementName")  # Replace 'YourElementName' with the name of the desired element
if desired_element:
    extracted_content = str(desired_element.prettify())

# Step 2: Construct new XML using the structure from the second XML file

with open("ecam_vba_ribbon.xml", "r") as file:
    content = file.read()
    soup2 = BeautifulSoup(content, 'lxml-xml')

# Here, you'd use the previous code for modifying and constructing the new XML.
# I'm not repeating it for brevity, but it involves creating new tags based on your requirements.


# Step 3: Combine and save the result

combined_content = extracted_content + "\n" + str(soup2)

with open("manifest_bs4.xml", "w") as f:
    f.write(combined_content)

