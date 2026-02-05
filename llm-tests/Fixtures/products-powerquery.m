// Power Query M code - Creates a sample products table
// This fixture creates data inline (no external file dependency)
let
    Source = #table(
        type table [
            ProductID = Int64.Type,
            ProductName = Text.Type,
            Category = Text.Type,
            Price = Decimal.Type,
            Rating = Decimal.Type
        ],
        {
            {1001, "Wireless Headphones", "Electronics", 79.99, 4.5},
            {1002, "Laptop Stand", "Accessories", 49.99, 4.2},
            {1003, "USB-C Hub", "Electronics", 39.99, 4.7},
            {1004, "Mechanical Keyboard", "Electronics", 129.99, 4.8},
            {1005, "Monitor Light Bar", "Accessories", 59.99, 4.3},
            {1006, "Webcam HD", "Electronics", 89.99, 4.1},
            {1007, "Desk Organizer", "Office", 24.99, 4.0},
            {1008, "Cable Management Kit", "Accessories", 19.99, 4.4}
        }
    )
in
    Source
