import time

def time_cmp(first_time, second_time):
    print( first_time)
    print(time.strftime("%H%M%S", first_time))
    print(time.strftime("%H%M%S", second_time))
    return int(time.strftime("%H%M%S", first_time)) - int(time.strftime("%H%M%S", second_time))

print(time_cmp(time.gmtime(), time.strptime("09:35:10", "%H:%M:%S")))