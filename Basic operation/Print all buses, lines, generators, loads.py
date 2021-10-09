import powerfactory as pf

app = pf.GetApplication()
app.ClearOutputWindow()

# Create dictionary of buses
bus_dict = {}
buses = app.GetCalcRelevantObjects('*.ElmTerm')
for bus in buses:
  bus_dict[bus.loc_name] = bus
  app.PrintPlain("bus %s" % (bus))

# Create dictionary of lines 
line_dict = {}
lines = app.GetCalcRelevantObjects('*.ElmLne')
for line in lines:
  line_dict[line.loc_name] = line
  app.PrintPlain("Line %s" % (line))    
    
# Create dictionary of synchronous generators 
gen_dict = {}
gens = app.GetCalcRelevantObjects('*.ElmSym')
for gen in gens:
  gen_dict[gen.loc_name] = gen
  app.PrintPlain("gen %s" % (gen))

# Loop through generator dictionary and print key and generator object
#for gen_key in gen_dict.keys():
#  app.PrintPlain(gen_key)
#  app.PrintPlain(gen_dict[gen_key])
  
# Create dictionary of loads 
load_dict = {}
loads = app.GetCalcRelevantObjects('*.ElmLod')
for load in loads:
  load_dict[load.loc_name] = load
  app.PrintPlain("Load %s" % (load))
  
# Loop through generator dictionary and print key and generator object
for bus_key in bus_dict.keys():
  app.PrintPlain(bus_key)
  app.PrintPlain(bus_dict[bus_key])