namespace OperateExcel.Job;

public sealed class DailyReportProfileOptions
{
    public List<DailyReportPersonOptions> People { get; set; } = [];
    public List<DailyReportStoreOptions> Stores { get; set; } = [];
    public List<DailyReportSkuOwnerOptions> SkuOwners { get; set; } = [];
}

public sealed class DailyReportPersonOptions
{
    public string? Name { get; set; }
    public double PaymentMonthlyBudget { get; set; }
}

public sealed class DailyReportStoreOptions
{
    public string? Name { get; set; }
    public List<string> People { get; set; } = [];
    public List<string> SourceDirectoryNames { get; set; } = [];
    public List<string> MappingSheetNameCandidates { get; set; } = [];
}

public sealed class DailyReportSkuOwnerOptions
{
    public string? OwnerCode { get; set; }
    public string? Name { get; set; }
    public string? AccountName { get; set; }
}

internal sealed class DailyReportProfile
{
    private DailyReportProfile(
        IReadOnlyList<string> people,
        IReadOnlyList<string> stores,
        IReadOnlyDictionary<string, double> paymentMonthlyBudgetByPerson,
        IReadOnlyDictionary<string, IReadOnlyList<string>> storePeople,
        IReadOnlyDictionary<string, string> sourceDirectoryAccountNames,
        IReadOnlyList<string> sourceDirectoryReadOrder,
        IReadOnlyDictionary<string, IReadOnlyList<string>> mappingSheetNameCandidatesByStore,
        IReadOnlyList<DailyReportSkuOwner> skuOwners,
        IReadOnlyDictionary<string, string> skuOwnerNamesByCode)
    {
        People = people;
        Stores = stores;
        PaymentMonthlyBudgetByPerson = paymentMonthlyBudgetByPerson;
        StorePeople = storePeople;
        SourceDirectoryAccountNames = sourceDirectoryAccountNames;
        SourceDirectoryReadOrder = sourceDirectoryReadOrder;
        MappingSheetNameCandidatesByStore = mappingSheetNameCandidatesByStore;
        SkuOwners = skuOwners;
        SkuOwnerNamesByCode = skuOwnerNamesByCode;
    }

    public IReadOnlyList<string> People { get; }
    public IReadOnlyList<string> Stores { get; }
    public IReadOnlyDictionary<string, double> PaymentMonthlyBudgetByPerson { get; }
    public IReadOnlyDictionary<string, IReadOnlyList<string>> StorePeople { get; }
    public IReadOnlyDictionary<string, string> SourceDirectoryAccountNames { get; }
    public IReadOnlyList<string> SourceDirectoryReadOrder { get; }
    public IReadOnlyDictionary<string, IReadOnlyList<string>> MappingSheetNameCandidatesByStore { get; }
    public IReadOnlyList<DailyReportSkuOwner> SkuOwners { get; }
    public IReadOnlyDictionary<string, string> SkuOwnerNamesByCode { get; }

    public static DailyReportProfile FromOptions(DailyReportProfileOptions options)
    {
        var people = new List<string>();
        var peopleSet = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var paymentMonthlyBudgetByPerson = new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase);

        foreach (var person in options.People)
        {
            var personName = NormalizeRequired(person.Name, "DailyReportProfile:People[].Name");
            if (!peopleSet.Add(personName))
            {
                throw new InvalidOperationException($"Duplicate person name in DailyReportProfile:People: {personName}");
            }

            people.Add(personName);
            paymentMonthlyBudgetByPerson[personName] = person.PaymentMonthlyBudget;
        }

        if (people.Count == 0)
        {
            throw new InvalidOperationException("DailyReportProfile:People must contain at least one person.");
        }

        var stores = new List<string>();
        var storeSet = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var storePeople = new Dictionary<string, IReadOnlyList<string>>(StringComparer.OrdinalIgnoreCase);
        var sourceDirectoryAccountNames = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var sourceDirectoryReadOrder = new List<string>();
        var sourceDirectoryReadOrderSet = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var mappingSheetNameCandidatesByStore = new Dictionary<string, IReadOnlyList<string>>(StringComparer.OrdinalIgnoreCase);

        foreach (var store in options.Stores)
        {
            var storeName = NormalizeRequired(store.Name, "DailyReportProfile:Stores[].Name");
            if (!storeSet.Add(storeName))
            {
                throw new InvalidOperationException($"Duplicate store name in DailyReportProfile:Stores: {storeName}");
            }

            stores.Add(storeName);

            var configuredPeople = TrimDistinct(store.People).ToList();
            if (configuredPeople.Count == 0)
            {
                configuredPeople.AddRange(people);
            }

            foreach (var personName in configuredPeople)
            {
                if (!peopleSet.Contains(personName))
                {
                    throw new InvalidOperationException(
                        $"Store '{storeName}' references person '{personName}', but that person is not defined in DailyReportProfile:People.");
                }
            }

            storePeople[storeName] = configuredPeople;

            var mappingCandidates = TrimDistinct(store.MappingSheetNameCandidates).ToList();
            if (mappingCandidates.Count == 0)
            {
                throw new InvalidOperationException(
                    $"Store '{storeName}' must define at least one MappingSheetNameCandidates value.");
            }

            mappingSheetNameCandidatesByStore[storeName] = mappingCandidates;

            var sourceDirectoryNames = TrimDistinct(store.SourceDirectoryNames).ToList();
            if (sourceDirectoryNames.Count == 0)
            {
                sourceDirectoryNames.Add(storeName);
            }

            foreach (var sourceDirectoryName in sourceDirectoryNames)
            {
                AddSourceDirectoryAccountName(sourceDirectoryAccountNames, sourceDirectoryName, storeName);
                if (sourceDirectoryReadOrderSet.Add(sourceDirectoryName))
                {
                    sourceDirectoryReadOrder.Add(sourceDirectoryName);
                }
            }

            AddSourceDirectoryAccountName(sourceDirectoryAccountNames, storeName, storeName);
            if (sourceDirectoryReadOrderSet.Add(storeName))
            {
                sourceDirectoryReadOrder.Add(storeName);
            }
        }

        if (stores.Count == 0)
        {
            throw new InvalidOperationException("DailyReportProfile:Stores must contain at least one store.");
        }

        var skuOwners = new List<DailyReportSkuOwner>();
        var skuOwnerNamesByCode = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var owner in options.SkuOwners)
        {
            var ownerCode = NormalizeRequired(owner.OwnerCode, "DailyReportProfile:SkuOwners[].OwnerCode");
            var personName = NormalizeRequired(owner.Name, "DailyReportProfile:SkuOwners[].Name");
            var accountName = owner.AccountName?.Trim();

            if (!peopleSet.Contains(personName))
            {
                throw new InvalidOperationException(
                    $"Sku owner '{ownerCode}' references person '{personName}', but that person is not defined in DailyReportProfile:People.");
            }

            var normalizedOwnerCode = NormalizePersonName(ownerCode);
            if (!skuOwnerNamesByCode.TryAdd(normalizedOwnerCode, personName))
            {
                throw new InvalidOperationException($"Duplicate sku owner code in DailyReportProfile:SkuOwners: {ownerCode}");
            }

            skuOwners.Add(new DailyReportSkuOwner(ownerCode, personName, string.IsNullOrWhiteSpace(accountName) ? null : accountName));
        }

        if (skuOwners.Count == 0)
        {
            throw new InvalidOperationException("DailyReportProfile:SkuOwners must contain at least one owner mapping.");
        }

        return new DailyReportProfile(
            people,
            stores,
            paymentMonthlyBudgetByPerson,
            storePeople,
            sourceDirectoryAccountNames,
            sourceDirectoryReadOrder,
            mappingSheetNameCandidatesByStore,
            skuOwners,
            skuOwnerNamesByCode);
    }

    private static string NormalizeRequired(string? value, string optionPath)
    {
        var normalized = value?.Trim() ?? string.Empty;
        if (normalized.Length == 0)
        {
            throw new InvalidOperationException($"{optionPath} cannot be empty.");
        }

        return normalized;
    }

    private static IEnumerable<string> TrimDistinct(IEnumerable<string?> values)
    {
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var value in values)
        {
            var normalized = value?.Trim() ?? string.Empty;
            if (normalized.Length > 0 && seen.Add(normalized))
            {
                yield return normalized;
            }
        }
    }

    private static void AddSourceDirectoryAccountName(
        IDictionary<string, string> accountNames,
        string sourceDirectoryName,
        string storeName)
    {
        if (accountNames.TryGetValue(sourceDirectoryName, out var existingStoreName)
            && !string.Equals(existingStoreName, storeName, StringComparison.OrdinalIgnoreCase))
        {
            throw new InvalidOperationException(
                $"Source directory name '{sourceDirectoryName}' is mapped to both '{existingStoreName}' and '{storeName}'.");
        }

        accountNames[sourceDirectoryName] = storeName;
    }

    private static string NormalizePersonName(string value)
    {
        return new string((value ?? string.Empty)
            .Trim()
            .Where(character => !char.IsWhiteSpace(character))
            .ToArray());
    }
}

internal sealed record DailyReportSkuOwner(string OwnerCode, string Name, string? AccountName);
