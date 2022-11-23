
var pascal = document.getElementById("pascal");
var méret = 10;

function faktor(n) {
    if (n === 0 || n === 1) {
        return 1;
    }
    else {
        return n * faktor(n - 1);
    }
}


for (var sor = 0; sor < méret; sor++) {

    var ujsor = document.createElement("div")
    ujsor.classList.add("sor")
    pascal.appendChild(ujsor)

    for (var oszlop = 0; oszlop <= sor; oszlop++) {
        var ujelem = document.createElement("div");

        ujelem.innerHTML = faktor(sor) / (faktor(oszlop) * faktor(sor - oszlop));

        ujelem.classList.add("elem");
        ujsor.appendChild(ujelem);
    }
}

