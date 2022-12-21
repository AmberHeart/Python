import sys

def fizz_buzz(limit):
    for i in range(1, limit+1):
        if i % 3 == 0:
            print('fizz',end="") 
        if i % 5 == 0:
            print('fizz',end="")
        if i % 3 and i % 5:
            print(i)
        else:
            print('fizz')

def main():
    fizz_buzz(int(sys.argv[1]))

if name == "__main__":
    main()