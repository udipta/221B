import pandas as pd

from dataset import BoM, Disti


def pn_quality_analytics(bom_data, disti_data):
    df = pd.DataFrame().assign(bom_pn=bom_data['Part Number'], bom_qty=bom_data.Quantity, dsti_pn='', dsti_qty='',
                               errorFlag=False)
    for i, row in df.iterrows():
        # check for BoM's PN in Disti
        disti = disti_data.loc[(disti_data['Part Number'] == row.bom_pn) & (disti_data.Quantity > 0), 'Quantity']
        disti_qty = disti.sum()
        if disti_qty:
            # check if Disti's quantity is more/less/equal of BoM's
            if disti_qty > row.bom_qty:  # greater
                val = abs(disti_qty - row.bom_qty)
                df.loc[i, ['dsti_pn', 'dsti_qty']] = [row.bom_pn, val]
                disti_data.loc[disti_data['Part Number'] == row.bom_pn, 'Quantity'] = val
            elif disti_qty < row.bom_qty:  # lesser
                df.loc[i, ['dsti_pn', 'dsti_qty']] = [row.bom_pn, disti_qty]
                disti_data.drop(disti.index, inplace=True)
            else:   # equal
                df.loc[i, ['dsti_pn', 'dsti_qty']] = [row.bom_pn, row.bom_qty]
                disti_data.drop(disti.index, inplace=True)
        else:
            df.loc[i, 'errorFlag'] = True

    # remain Disti data
    df = df.append(
        pd.DataFrame().assign(bom_pn='', bom_qty='', dsti_pn=disti_data['Part Number'], dsti_qty=disti_data.Quantity,
                              errorFlag=True))
    df.fillna('', inplace=True)

    return df


if __name__ == '__main__':
    try:
        bom_df = pd.DataFrame(BoM)
        disti_df = pd.DataFrame(Disti)
        data = pn_quality_analytics(bom_df, disti_df)
        print(data)
    except Exception as e:
        import traceback
        print(traceback.format_exc())
