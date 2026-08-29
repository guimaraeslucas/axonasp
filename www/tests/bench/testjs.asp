<%@LANGUAGE="JScript" %>
<%
var limit = 50000;
var primes = [];

var startTime = new Date().getTime();

for (var i = 2; i <= limit; i++) {
    var isPrime = true;
    for (var j = 2; j <= Math.floor(Math.sqrt(i)); j++) {
        if (i % j === 0) {
            isPrime = false;
            break;
        }
    }
    
    if (isPrime) {
        primes.push(i);
    }
}

var endTime = new Date().getTime();
var executionTime = endTime - startTime;

Console.log("JScript Execution Time: " + executionTime + " ms");
Console.log("Primes found: " + primes.length);
%>