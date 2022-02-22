/** a dumb function I wrote to solve a problem that didn't exist */
/*
function f(body,zip,element){
  // let body = template.getBody()
  // let zip = body.getChild(4)
  // let element = zip.copy()

  let w =[]
  let zipChilds = zip.getRow(0).getNumChildren()
  for (let i = 0;i<zipChilds;i++){
    w.push(zip.getColumnWidth(i))
  }

  if (zipChilds === 1){
    body.appendTable(element)
    .setColumnWidth(0,w[0])
  }
  else if (zipChilds === 2){
    body.appendTable(element)
    .setColumnWidth(0,w[0])
    .setColumnWidth(1,w[1])
  }
  else if (zipChilds === 3){
    body.appendTable(element)
    .setColumnWidth(0,w[0])
    .setColumnWidth(1,w[1])
    .setColumnWidth(2,w[2])
  }
  else if (zipChilds === 4){
    body.appendTable(element)
    .setColumnWidth(0,w[0])
    .setColumnWidth(1,w[1])
    .setColumnWidth(2,w[2])
    .setColumnWidth(3,w[3])
  } 
  else if (zipChilds === 5){
    body.appendTable(element)
    .setColumnWidth(0,w[0])
    .setColumnWidth(1,w[1])
    .setColumnWidth(2,w[2])
    .setColumnWidth(3,w[3])
    .setColumnWidth(4,w[4])
  }

}
*/