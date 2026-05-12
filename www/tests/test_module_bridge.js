import { getSeed } from "./test_module_math.js";

export function render(label) {
    Response.Write("bridge");
    return label + ":" + getSeed();
}
