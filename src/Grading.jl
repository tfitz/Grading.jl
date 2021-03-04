module Grading

using DataFrames
using XLSX

struct category
    Name::String
    weight::Real
    numberToDrop::Int
    list::Array{String,1}

    # check to make sure the item's weight is not negative or above 1
    category(Name,weight,numberToDrop,list) = (weight < 0) | (weight > 1) ? error("Weighting factor not properly defined") : new(Name,weight,numberToDrop, list)
end

function checkTotalWeights(list::Array{category,1})
    total = 0.
    for item in list
        total += item.weight
    end

    if total ≈ 1
        return true
    else
        return false
    end
end

function normalize_row(row, item::category)

    output = Float64[]
    
    for str in item.list
        idx = Symbol(str)
        idx_maxPt = Symbol( str * " - Max Points" )
           
        if ismissing( row[idx] )
            val = 0
        else
            val = row[idx][1]
        end
    
        append!(output, val/row[idx_maxPt][1] )
        # println( val/row[idx_maxPt][1] )
    end

    return output
end

function computeDropMean(data::Array{Float64}, lows_to_drop::Int)

    n = length(data)
    lows = sort(data)[1:lows_to_drop]
    total = sum(data)
    
    dropmean = ( total - sum(lows) )/( n - lows_to_drop )
    return dropmean
end

function computeScore(row, list::Array{category,1} )

    if !checkTotalWeights(list)
        error("Sum of weights is not 1.0")
    end

    score = 0.
    for item in list
        data = normalize_row(row, item)
        local_score = computeDropMean(data, item.numberToDrop )
        score += item.weight*local_score
    end

    return score

end

function computeEachFinalScore!(data, list::Array{category,1}; ColumnName="Final Score" )

    if !checkTotalWeights(list)
        error("Sum of weights is not 1.0")
    end

    data[!, Symbol(ColumnName)] .= [ computeScore(row, list::Array{category,1} ) for row in eachrow(data) ]

end

function computeEachCategoryScore!(data::DataFrames.DataFrame, list::Array{category,1} )
    
    for item in list
         data[!, Symbol(item.Name)] .= [ computeDropMean(normalize_row(row, item), item.numberToDrop ) for row in eachrow(data) ]
    end

end

function importGradescopeXSLX(infilename::String; SheetName = "Course Grades")
    data0 = XLSX.readxlsx(infilename)[SheetName]
    data = XLSX.eachtablerow(data0) |> DataFrames.DataFrame
    return data
end

function exportGradeSheet(outfilename::String, data; SheetName = "Course Grades")

    XLSX.writetable(outfilename, data, overwrite=true, sheetname=SheetName)

end

end