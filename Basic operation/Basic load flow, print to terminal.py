if __name__ == "__main__":
  
  #connect to PowerFactory
  import powerfactory as pf
  
  app = pf.GetApplication()
  app.ClearOutputWindow()
  
  if app is None:
    raise Exception("getting PowerFactory application failed")
  
  #print to PowerFactory output window
  app.PrintInfo("Python Script started..")

  #get active project
  prj = app.GetActiveProject()
  if prj is None:
    raise Exception("No project activated. Python Script stopped.")

  #retrieve load-flow object
  ldf = app.GetFromStudyCase("ComLdf")

  #force balanced load flow
  ldf.iopt_net = 0

  #execute load flow
  ldf.Execute()

  #collect all relevant terminals
  app.PrintInfo("Collecting all calculation relevant terminals..")
  terminals = app.GetCalcRelevantObjects("*.ElmTerm")
  
  if not terminals:
    raise Exception("No calculation relevant terminals found")
  
  app.PrintPlain("Number of terminals found: %d" % len(terminals))
  
  for terminal in terminals:
    voltage = terminal.GetAttribute("m:u")
    app.PrintPlain("Voltage at terminal %s is %f p.u." % (terminal , voltage))

  #print to PowerFactory output window
  app.PrintInfo("Python Script ended.")