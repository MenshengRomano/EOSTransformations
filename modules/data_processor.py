import pandas as pd


def apply_transformation(transformation, value, context=None):
    try:
        if pd.isna(value):
            return value
        if transformation == "Value = 0":
            return 0
        elif transformation == "UNMAPPED":
            return None
        elif transformation == "Value = Value":
            return value
        elif transformation == "LookupUnit":
            tCode3 = context['tCode3']
            if value is not None:
                code3_value = str(value).split(" ")[0]  # Get the leading number
                lookup_result = tCode3[tCode3['ID'] == code3_value]
                if not lookup_result.empty:
                    return lookup_result['Aggregate unit'].values[0]
            return "LS"
        elif transformation == 'SUBSTITUTE(Value, " || "," |  ")':
            return value.replace(" || ", " |  ")
        elif transformation == 'SUBSTITUTE(Value, " || "," | ")':
            return value.replace(" || ", " | ")
        elif transformation == 'SUBSTITUTE("Value", " || ","") & " | "':
            return value.replace(" || ", "") + " | "
        elif transformation == "Value = 'GT'":
            return "GT | Grand total"
        else:
            return value  # Default: no transformation
    except Exception as e:
        print(f"Error applying transformation '{transformation}' on value '{value}': {e}")
        return value  # Return original value on error


def process_data(source_df, tables, item_ws):
    mapping_df = tables["tMapping"]

    # Filter out rows where Target Column is blank or null
    mapping_df = mapping_df[mapping_df["Target Column"].notna() & (mapping_df["Target Column"] != '')]

    # Check presence of Source Column and Target Column in source_df and item_ws
    mapping_df['Source Column Found'] = mapping_df['Source Column'].apply(
        lambda x: x in source_df.columns if x != '{N/A}' else False)
    item_headers = [cell.value for cell in item_ws[1]]
    mapping_df['Target Column Found'] = mapping_df['Target Column'].apply(lambda x: x in item_headers)

    # Initialize item_df with columns specified in the mapping table
    item_df = pd.DataFrame(columns=mapping_df["Target Column"].unique())

    tCode3 = tables.get("tCode3", pd.DataFrame())
    context = {'tCode3': tCode3}

    filtered_source_df = source_df[source_df["Cost Code 3 (Review)"].notna()]
    print(f"Filtered source_df with {len(filtered_source_df)} rows")

    # Prepare data for the "Item" sheet based on the mapping table
    item_rows = []
    for idx, source_row in filtered_source_df.iterrows():
        item_row = {}
        for _, row in mapping_df.iterrows():
            source_column = row["Source Column"]
            target_column = row["Target Column"]
            transformation = row["Transformation"]

            print(f"Processing mapping from {source_column} to {target_column} with transformation: {transformation}")

            # Apply transformation if specified
            try:
                if pd.notna(transformation) and transformation != '{N/A}':
                    if source_column != '{N/A}':
                        item_row[target_column] = apply_transformation(transformation, source_row[source_column],
                                                                       context)
                    else:
                        # Handle special case for QuantityUnit
                        if target_column == "QuantityUnit" and transformation == "LookupUnit":
                            cost_code_value = source_row["Cost Code 3 (Review)"]
                            item_row[target_column] = apply_transformation(transformation, cost_code_value, context)
                        else:
                            item_row[target_column] = apply_transformation(transformation, None, context)
                else:
                    if source_column == '{N/A}':
                        item_row[target_column] = apply_transformation(transformation, None, context)
                    else:
                        item_row[target_column] = source_row[source_column]
            except Exception as e:
                print(f"Error applying transformation '{transformation}' on column '{source_column}': {e}")
                item_row[target_column] = source_row[source_column]  # Use original value on error

        item_rows.append(item_row)

    # Convert the list of item rows to a DataFrame
    item_df = pd.DataFrame.from_records(item_rows, columns=item_df.columns)

    # Ensure that the QuantityUnit column is correctly populated
    if 'QuantityUnit' in item_df.columns:
        item_df['QuantityUnit'] = item_df['QuantityUnit'].fillna("LS")

    # Handle renaming and reordering columns
    if 'Quantity' in item_df.columns and 'QuantityUnit' in item_df.columns:
        cols = list(item_df.columns)
        cols.insert(cols.index('Quantity') + 1, cols.pop(cols.index('QuantityUnit')))
        item_df = item_df[cols]

    # Ensure values for hardcoded transformations
    if 'RowLevel' in item_df.columns:
        item_df['RowLevel'] = item_df['RowLevel'].fillna(0)
    if 'Group by Grand total' in item_df.columns:
        item_df['Group by Grand total'] = item_df['Group by Grand total'].fillna('GT | Grand total')

    item_df = item_df.fillna('')
    column_mapping = {col: item_headers.index(col) + 1 for col in item_df.columns if col in item_headers}

    # Write data to the "Item" worksheet
    for r_idx, row in enumerate(item_df.itertuples(index=False, name=None), 2):
        for col, value in zip(item_df.columns, row):
            if col in column_mapping:
                c_idx = column_mapping[col]
                item_ws.cell(row=r_idx, column=c_idx, value=value)
        # Handle QuantityUnit specifically
        if 'Quantity' in column_mapping and 'QuantityUnit' in item_df.columns:
            quantity_col_idx = column_mapping['Quantity']
            item_ws.cell(row=r_idx, column=quantity_col_idx + 1, value=row[item_df.columns.get_loc('QuantityUnit')])

    return item_df, mapping_df
