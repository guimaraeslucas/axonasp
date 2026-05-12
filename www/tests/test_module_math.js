export var seed = 5;

export function add(a, b) {
    return a + b;
}

export function getSeed() {
    return seed;
}

export function bump() {
    seed = seed + 1;
    return seed;
}
