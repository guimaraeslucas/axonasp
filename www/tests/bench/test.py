import time
import math

def benchmark_primes(limit):
    start_time = time.time()
    primes = []
    
    for num in range(2, limit + 1):
        is_prime = True
        for i in range(2, int(math.sqrt(num)) + 1):
            if num % i == 0:
                is_prime = False
                break
        if is_prime:
            primes.append(num)
            
    end_time = time.time()
    execution_time = (end_time - start_time) * 1000 # Conversão para milissegundos
    
    print(f"Python Execution Time: {execution_time:.2f} ms")
    print(f"Primes found: {len(primes)}")

if __name__ == "__main__":
    benchmark_primes(50000)