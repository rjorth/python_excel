import pdftables_api 

c = pdftables_api.Client('') #api key
c.xlsx('HAY DKK.pdf', 'output') #replace c.xlsx with c.csv to convert to CSV