function transpose(matrix) {
  return matrix[0].map((col, i) => matrix.map(row => row[i]));
}

function NAMINGCONVENTION(text,convention){
  if(convention = "CapitalizeFirstLetter"){
      return text.charAt(0).toUpperCase() + text.slice(1);
  }
}