import pandas as pd

main =  pd.read_csv("CIS campaign list 4.csv")

numberOfBatches = int(input("Enter number of batches: "))

batchSize = int(len(main) / numberOfBatches)

print(len(main))
iprev, ipres = 0, batchSize
for i in range(1, numberOfBatches + 1):
    print(ipres)
    batch = main.iloc[iprev:,:] if ipres >= len(main) - 1 else main.iloc[iprev:ipres,:]
    iprev = ipres
    ipres += batchSize
    batch.to_csv(rf"batch {i}.csv", index = False)