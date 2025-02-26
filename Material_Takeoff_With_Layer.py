"""Revit Dynamo script for extracting material information from selected elements.

This script analyzes the compound structure of selected Revit elements and extracts
detailed material information for each layer. It works with any element type that has
a compound structure (e.g., walls, floors, roofs, etc.).

Inputs:
    IN[0]: Selected Revit elements (single element or list of elements)

Outputs:
    OUT: List of dictionaries containing material information with the following keys:
        - Family Type: Name of the family
        - Type Name: Name of the type
        - Category: Category of the element
        - Material: Name of the material
        - Layer Number: Index of the layer (1-based)

Exceptions:
    ValueError: If no elements are selected
    Various Revit API exceptions are caught and logged

Author: Tamrat Mengistu  
Last Modified: 2026-26-02
"""

import clr
clr.AddReference('RevitServices')
from RevitServices.Persistence import DocumentManager
clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import *
from System.Collections.Generic import List
from System import InvalidOperationException

def get_parameter_value(element, param_name, default=""):
    """Safely get parameter value from an element.
    
    Args:
        element: Revit element
        param_name: Built-in parameter to retrieve
        default: Default value if parameter is not found or invalid
    
    Returns:
        Parameter value as string or default value
    """
    try:
        param = element.get_Parameter(param_name)
        if param and param.HasValue:
            return param.AsString()
        return default
    except Exception as e:
        print(f"Error getting parameter {param_name}: {str(e)}")
        return default

# Get current document
try:
    doc = DocumentManager.Instance.CurrentDBDocument
    if not doc:
        raise InvalidOperationException("Could not access current document")
except Exception as e:
    raise InvalidOperationException(f"Failed to get document: {str(e)}")

# Process input selection
try:
    selection = IN[0]
    if selection is None:
        raise ValueError("Input selection is null")
        
    elements = selection if isinstance(selection, list) else [selection]
    
    if not elements:
        raise ValueError("Please select some elements first")
except Exception as e:
    raise ValueError(f"Error processing input selection: {str(e)}")

# Get element types from selection
element_types = []
for element in UnwrapElement(elements):
    try:
        type_id = element.GetTypeId()
        if type_id and not type_id.IsNull:
            element_type = doc.GetElement(type_id)
            if element_type and element_type not in element_types:
                element_types.append(element_type)
        else:
            print(f"Warning: Element has no valid type ID")
    except Exception as e:
        print(f"Error getting element type: {str(e)}")
        continue

material_data = []

# Parse compound structures
for element_type in element_types:
    try:
        # Get basic element information
        family_name = get_parameter_value(element_type, BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM, "Unknown Family")
        type_name = get_parameter_value(element_type, BuiltInParameter.SYMBOL_NAME_PARAM, "Unknown Type")
        category_name = element_type.Category.Name if element_type.Category else "Unknown Category"
        
        compound_structure = element_type.GetCompoundStructure()
        if not compound_structure:
            print(f"Warning: No compound structure found for {type_name}")
            continue
            
        layers = compound_structure.GetLayers()
        if not layers:
            print(f"Warning: No layers found in compound structure for {type_name}")
            continue
            
        for idx, layer in enumerate(layers):
            try:
                material = doc.GetElement(layer.MaterialId)
                if material:
                    material_data.append({
                        "Family Type": family_name,
                        "Type Name": type_name,
                        "Category": category_name,
                        "Material": material.Name,
                        "Layer Number": idx + 1
                    })
                else:
                    print(f"Warning: Material not found for layer in {type_name}, Layer Index: {idx+1}")
            except Exception as e:
                print(f"Error processing layer {idx+1} in {type_name}: {str(e)}")
                continue
    except Exception as e:
        print(f"Error processing element type {type_name}: {str(e)}")
        continue

OUT = material_data
