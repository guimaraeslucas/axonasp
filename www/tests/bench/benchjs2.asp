<%@ Language="JScript" %>
<%
// AxonASP JScript Engine Benchmark Suite
// Architecture: Modular execution with timing telemetry via console.log

var BASE_ITERATIONS = 1000000;

function runBenchmark(testName, testFunction) {
    var startTime = new Date().getTime();
    testFunction();
    var endTime = new Date().getTime();
    var duration = endTime - startTime;
    
    console.log("[AxonASP Benchmark] " + testName + " completed in " + duration + " ms.");
    return duration;
}

function benchMath() {
    var result = 0;
    var iterations = BASE_ITERATIONS * 5; // 5M iterations
    for (var i = 0; i < iterations; i++) {
        result = (Math.sin(i) * Math.cos(i)) + Math.sqrt(i) + Math.pow(i, 2);
    }
}

function benchStrings() {
    var str = "";
    var iterations = BASE_ITERATIONS / 20; // 50k iterations to avoid exponential GC locks in older engines
    
    // Concatenation stress
    for (var i = 0; i < iterations; i++) {
        str += "a";
    }
    
    // RegExp and splitting stress
    var regex = /a/g;
    var matches = str.match(regex);
    var arr = str.split("");
}

function benchArrays() {
    var arr = [];
    var iterations = BASE_ITERATIONS / 10; // 100k iterations
    
    // Push stress
    for (var i = 0; i < iterations; i++) {
        arr.push(i);
    }
    
    // Mutation and sorting stress
    arr.reverse();
    arr.sort(function(a, b) {
        return a - b;
    });
}

function benchObjects() {
    var obj = {};
    var iterations = BASE_ITERATIONS; // 1M iterations
    
    // Property assignment stress
    for (var i = 0; i < iterations; i++) {
        obj["key_" + i] = i;
    }
    
    // Property access stress
    var readVal;
    for (var j = 0; j < iterations; j++) {
        readVal = obj["key_" + j];
    }
}

function benchLoops() {
    var count = 0;
    var iterations = BASE_ITERATIONS * 10; // 10M iterations
    
    // For loop stress
    for (var i = 0; i < iterations; i++) {
        count++;
    }
    
    // While loop stress
    var j = iterations;
    while (j > 0) {
        j--;
        count--;
    }
}

function executeSuite() {
    console.log("==========================================");
    console.log("Starting AxonASP JScript Benchmark Suite...");
    console.log("Base Iterations: " + BASE_ITERATIONS);
    console.log("==========================================");
    
    var totalStartTime = new Date().getTime();
    
    var t1 = runBenchmark("Math Operations", benchMath);
    var t2 = runBenchmark("String Manipulation", benchStrings);
    var t3 = runBenchmark("Array Operations", benchArrays);
    var t4 = runBenchmark("Object Instantiation & Access", benchObjects);
    var t5 = runBenchmark("Raw Iteration / Loops", benchLoops);
    
    var totalEndTime = new Date().getTime();
    var totalDuration = totalEndTime - totalStartTime;
    
    var totalMeasured = t1 + t2 + t3 + t4 + t5;
    var overhead = totalDuration - totalMeasured;
    
    console.log("==========================================");
    console.log("Total Benchmark Time : " + totalDuration + " ms");
    console.log("Engine Overhead/GC   : " + overhead + " ms");
    console.log("==========================================");
}

// Execute the benchmark suite
executeSuite();
%>