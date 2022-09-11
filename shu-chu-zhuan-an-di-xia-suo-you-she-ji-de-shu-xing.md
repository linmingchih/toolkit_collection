---
description: Output Properties of All HFSS Designs in the Project to CSV for Comparison
---

# 輸出專案底下所有設計的屬性

<figure><img src=".gitbook/assets/image (1).png" alt=""><figcaption></figcaption></figure>

<details>

<summary>output_properties_table</summary>

{% code lineNumbers="true" %}
```python
from pyaedt import Hfss

hfss = Hfss(specified_version='2022.1')

properties = {}
prop_list = []
for design in hfss.design_list:
    properties[design] = {}
    x = Hfss(projectname=hfss.project_name, designname=design)
    props = x.odesign.GetProperties('LocalVariableTab', 'LocalVariables')
    for i in props:
        if i not in prop_list:
            prop_list.append(i)
        properties[design][i] = x.odesign.GetPropertyValue('LocalVariableTab', 'LocalVariables', i)
    
prop_list.sort()

with open('d:/demo/tabls.csv', 'w') as f:
    f.writelines(','.join(hfss.design_list))
    for p in prop_list:
        row = [p]
        for design in hfss.design_list:
            try:
                row.append(properties[design][p])
            except:
                row.append('NaN')
        f.writelines(','.join(row) + '\n')
```
{% endcode %}

</details>

