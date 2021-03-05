import matplotlib.pyplot as plt

def get_data():
    with open('data.csv') as f:
        return f.read()

    return None

def main():
    data = get_data()
    rows = data.split('\n')
    rows = rows[1:]

    shoe_sizes = {}
    for row in rows:
        cols = row.split(',')
        if (len(cols) < 2):
            continue
        
        name = cols[1]
        words = name.split(' ')
        word_size = words[-1]
        
        try:
            # to ensure it is a float
            # size is float therefore due to way machines store
            # decimals cannot guarantee 4.5 == 4.50000...
            size = float(word_size)
            
            # string key is better
            # key = str(size)
            # legacy
            key = size
            
            if key not in shoe_sizes:
                shoe_sizes[key] = 0
            
            shoe_sizes[key] += 1
        except ValueError:
            print('error has occured, size is not a float')
            continue

        # try:
            # print('word size is %s' % (word_size))
            # size = float(word_size)
            # shoe_sizes.append(size)

            # print('size is %f' % (size))
        # except ValueError:
            # print('error has occured, size is not a float')
        #    continue

    shoe_sizes_length = len(shoe_sizes)
    x_axis = range(shoe_sizes_length)
    
    print(shoe_sizes)
    
    keys = list(shoe_sizes.keys())
    # sort keys
    keys.sort()
    
    # if keys are sorted therefore iterating through keys
    # values (i.e heights) are sorted too
    heights = []
    
    for key in keys:
        value = shoe_sizes[key]
        heights.append(value)
    
    plt.bar(x_axis, heights, align='center')
    plt.xticks(x_axis, keys)
    # plt.bar(x_axis, list(shoe_sizes.values()), align='center')
    # plt.xticks(x_axis, list(shoe_sizes.keys()))
    
    # plt.plot(shoe_sizes)
    # print(len(shoe_sizes))
    # print(shoe_sizes)
    # plt.ylabel('shoe sizes')
    # print('showing graph')
    plt.show()

        # print('name is %s' % (name))
    # print('data is %s' % (data))
    # print('rows', rows)

if __name__ == '__main__':
    main()
