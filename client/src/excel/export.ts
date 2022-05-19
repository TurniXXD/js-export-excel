import { ETableType, ESummaryTabType, EModule } from './constants'

import XLSX from 'xlsx'

//import { ExportExcel } from '@/client/excel/summary/exportDefinition'

//const { t } = useTranslation()

// <Button
// label={t('summary.exportExcel')}
// onClick={() => ExportExcel(TableOverviewSelectVirtual, ESummaryTabType.Table, t, Nursery, customerModules)}
// >

export const ExportExcel = (array: any, type: ESummaryTabType, t: any, nurseries?: any, customerModules?: any) => {
  if (array.length === 0) return

  const measureContext = document.createElement('canvas').getContext('2d')!
  measureContext.font = '12px arial'

  // Calculate string width in pixels
  const textToPixels = (text?: string | number | null) =>
    text ? Math.ceil(measureContext.measureText(text?.toString() || '').width + 10) : 0

  const updateColumnWidth = <T extends Record<string, number>, K extends keyof T>(
    value: T[K],
    widths: T,
    column: K
  ) => {
    value > widths[column] && (widths[column] = value)
  }

  const editTabsData = (centerTabActive: boolean, identifierLabel: string, subIdentifierLabel: string) => {
    interface SummaryJson {
      identifier: string
      subIdentifier: string
      areaAndLevel: string
      section: string
      type?: string
      age?: string
      seed?: number
      material?: number
      plants?: number
      plantParts?: number
    }

    const summary: Array<SummaryJson> = []
    const summaryColumnsWidth: Required<Record<keyof SummaryJson, number>> = {
      identifier: textToPixels(identifierLabel),
      subIdentifier: textToPixels(subIdentifierLabel),
      areaAndLevel: textToPixels(t('summary.headers.areaAndLevel')),
      section: textToPixels(t('summary.headers.section')),
      type: textToPixels(t('summary.headers.type')),
      age: textToPixels(t('summary.headers.age')),
      seed: textToPixels(t('summary.headers.seed')),
      material: textToPixels(t('summary.headers.material')),
      plants: textToPixels(t('summary.headers.plants')),
      plantParts: textToPixels(t('summary.headers.plantParts')),
    }

    array?.map((MaterialOverviewsVirtualItem: any) => {
      const center = MaterialOverviewsVirtualItem.Center?.name,
        woodSpecy = `${MaterialOverviewsVirtualItem.ApprovedArea.WoodSpecy?.code} - ${MaterialOverviewsVirtualItem.ApprovedArea.WoodSpecy?.name}`,
        areaAndLevel = `${MaterialOverviewsVirtualItem.ApprovedArea.NaturalWoodArea?.name} - ${MaterialOverviewsVirtualItem.ApprovedArea.ForestVegetativeLevel?.name}`,
        section = `${MaterialOverviewsVirtualItem?.order.toString().padStart(4, '0')} / ${
          MaterialOverviewsVirtualItem?.year
        }`
      let identifier: string, subIdentifier: string

      centerTabActive ? (identifier = center) : (identifier = woodSpecy)
      centerTabActive ? (subIdentifier = woodSpecy) : (subIdentifier = center)

      MaterialOverviewsVirtualItem.SectionStateSeed?.length === 0 &&
        MaterialOverviewsVirtualItem.SectionStatePlant?.length === 0 &&
        MaterialOverviewsVirtualItem.SectionStatePlantPart?.length === 0 &&
        summary.push({
          identifier,
          subIdentifier,
          areaAndLevel,
          section,
        })

      const sectionRowsPush = (
        sectionRowsArray: any,
        plantMaterialType: string,
        age: string,
        seed: number,
        material: number,
        plants: number,
        plantParts: number
      ) => {
        sectionRowsArray.push({
          identifier,
          subIdentifier,
          areaAndLevel,
          section,
          plantMaterialType,
          age,
          seed,
          material,
          plants,
          plantParts,
        })

        updateColumnWidth(textToPixels(plantMaterialType), summaryColumnsWidth, 'type')
        updateColumnWidth(textToPixels(age), summaryColumnsWidth, 'age')
        updateColumnWidth(textToPixels(seed), summaryColumnsWidth, 'seed')
        updateColumnWidth(textToPixels(material), summaryColumnsWidth, 'material')
        updateColumnWidth(textToPixels(plants), summaryColumnsWidth, 'plants')
        updateColumnWidth(textToPixels(plantParts), summaryColumnsWidth, 'plantParts')
      }

      MaterialOverviewsVirtualItem.SectionStateSeed?.length &&
        MaterialOverviewsVirtualItem.SectionStateSeed.map((SectionStateSeedItem: any) => {
          const plantMaterialType = t('summary.plantMaterial.SEED'),
            age = SectionStateSeedItem.age || '?',
            seed = SectionStateSeedItem?.amountSeed,
            material = SectionStateSeedItem?.amountMaterial,
            plants = 0,
            plantParts = 0

          sectionRowsPush(summary, plantMaterialType, age, seed, material, plants, plantParts)
        })

      MaterialOverviewsVirtualItem.SectionStatePlant?.length &&
        MaterialOverviewsVirtualItem.SectionStatePlant.map((SectionStatePlantItem: any) => {
          const plantMaterialType = t('summary.plantMaterial.PLANT'),
            age = SectionStatePlantItem?.age || '?',
            seed = 0,
            material = 0,
            plants = SectionStatePlantItem?.amount,
            plantParts = 0

          sectionRowsPush(summary, plantMaterialType, age, seed, material, plants, plantParts)
        })

      MaterialOverviewsVirtualItem.SectionStatePlantPart?.length &&
        MaterialOverviewsVirtualItem.SectionStatePlantPart.map((SectionStatePlantPartItem: any) => {
          const plantMaterialType = t('summary.plantMaterial.PLANT_PARTS'),
            age = SectionStatePlantPartItem?.age || '?',
            seed = 0,
            material = 0,
            plants = 0,
            plantParts = SectionStatePlantPartItem?.amount

          sectionRowsPush(summary, plantMaterialType, age, seed, material, plants, plantParts)
        })

      updateColumnWidth(textToPixels(identifier), summaryColumnsWidth, 'identifier')
      updateColumnWidth(textToPixels(subIdentifier), summaryColumnsWidth, 'subIdentifier')
      updateColumnWidth(textToPixels(section), summaryColumnsWidth, 'section')
      updateColumnWidth(textToPixels(areaAndLevel), summaryColumnsWidth, 'areaAndLevel')
    })

    // Create worksheet from updated array
    const ws = XLSX.utils.json_to_sheet([])

    // Set columns width
    ws['!cols'] = [
      { wpx: summaryColumnsWidth.identifier },
      { wpx: summaryColumnsWidth.subIdentifier },
      { wpx: summaryColumnsWidth.areaAndLevel },
      { wpx: summaryColumnsWidth.section },
      { wpx: summaryColumnsWidth.type },
      { wpx: summaryColumnsWidth.age },
      { wpx: summaryColumnsWidth.seed },
      { wpx: summaryColumnsWidth.material },
      { wpx: summaryColumnsWidth.plants },
      { wpx: summaryColumnsWidth.plantParts },
    ]

    const headers = [
      [
        identifierLabel,
        subIdentifierLabel,
        t('summary.headers.areaAndLevel'),
        t('summary.headers.section'),
        t('summary.headers.type'),
        t('summary.headers.age'),
        t('summary.headers.seed'),
        t('summary.headers.material'),
        t('summary.headers.plants'),
        t('summary.headers.plantParts'),
      ],
    ]

    XLSX.utils.sheet_add_aoa(ws, headers)
    XLSX.utils.sheet_add_json(ws, summary, { origin: 'A2', skipHeader: true })

    return ws
  }

  const editTableTabData = () => {
    interface SummaryJson {
      center: string
      tableCode: string
      type: string
      nursery?: string
      usedAcreage: string
      emptyAcreage: string
      acreage: number
      year: number
      occupancy: number
      // Optional SectionStatePlant row
      woodSpecy?: string
      naturalWoodArea?: string
      forestVegetativeLevel?: string
      age?: string
      amount?: string
      typeOperation?: string
      localization?: string
      areaAr?: string
      inventory?: boolean
      trimmed?: boolean
      section?: string
    }

    const isNursery = customerModules.includes(EModule.Nursery)
    let isSectionStatePlant = false

    const summary: Array<SummaryJson> = []
    const summaryColumnsWidth: Required<Record<keyof SummaryJson, number>> = {
      center: textToPixels(t('summary.headers.center')),
      tableCode: textToPixels(
        customerModules.includes(EModule.Nursery)
          ? t('section.tabs.operation.tableState.table')
          : t('section.tabs.operation.stock')
      ),
      type: textToPixels(t('section.tabs.operation.tableState.type')),
      ...(isNursery && { nursery: textToPixels(t('section.tabs.operation.tableState.nursery')) }),
      usedAcreage: textToPixels(t('section.tabs.operation.tableState.usedAreaAr')),
      emptyAcreage: textToPixels(t('section.tabs.operation.tableState.emptyAreaAr')),
      acreage: textToPixels(t('section.tabs.operation.tableState.acreage')),
      year: textToPixels(t('section.tabs.operation.tableState.year')),
      occupancy: textToPixels(t('section.tabs.operation.tableState.occupancy')),
      woodSpecy: textToPixels(t('section.tabs.operation.tableState.woodSpecies')),
      naturalWoodArea: textToPixels(t('section.tabs.operation.tableState.naturalWoodArea')),
      forestVegetativeLevel: textToPixels(t('section.tabs.operation.tableState.forestryVegetativeLevel')),
      age: textToPixels(t('section.tabs.operation.tableState.age')),
      amount: textToPixels(t('section.tabs.operation.tableState.amount')),
      typeOperation: textToPixels(t('section.tabs.operation.tableState.type')),
      localization: textToPixels(t('section.tabs.operation.tableState.localization')),
      areaAr: textToPixels(t('section.tabs.operation.tableState.usedAreaAr')),
      inventory: textToPixels(t('section.tabs.operation.tableState.inventory')),
      trimmed: textToPixels(t('section.tabs.operation.tableState.trimmed')),
      section: textToPixels(t('section.tabs.operation.tableState.section')),
    }

    array?.map((TableOverviewSelectVirtualCenter: any) => {
      const center = TableOverviewSelectVirtualCenter?.name

      TableOverviewSelectVirtualCenter.Nursery.map((TableOverviewSelectVirtualNursery: any) => {
        TableOverviewSelectVirtualNursery.Table.map((TableOverviewSelectVirtualItem: any) => {
          const isStock = TableOverviewSelectVirtualItem.type === ETableType.STOCK,
            areaAresUsed = TableOverviewSelectVirtualItem?.usedAcreage * 100 || 0,
            areaAresEmpty =
              (TableOverviewSelectVirtualItem?.acreage - TableOverviewSelectVirtualItem?.usedAcreage) * 100 || 0,
            isOversized = TableOverviewSelectVirtualItem?.acreage < TableOverviewSelectVirtualItem?.usedAcreage,
            tableCode = TableOverviewSelectVirtualItem.tableCode,
            type = TableOverviewSelectVirtualItem.type,
            nursery =
              isNursery && nurseries.find((nursery: any) => nursery.id === TableOverviewSelectVirtualNursery?.id).name,
            usedAcreage = !isStock ? areaAresUsed.toFixed(2) : '',
            emptyAcreage = !isStock ? areaAresEmpty.toFixed(2) : '',
            acreage = TableOverviewSelectVirtualItem.acreage,
            year = TableOverviewSelectVirtualItem.year,
            occupancy = !isStock
              ? isOversized
                ? 100
                : (100 * TableOverviewSelectVirtualItem?.usedAcreage) / TableOverviewSelectVirtualItem?.acreage || 0
              : ''

          let isSectionStatePlantItem: boolean

          TableOverviewSelectVirtualItem.SectionStatePlant?.length > 0
            ? (isSectionStatePlantItem = true)
            : (isSectionStatePlantItem = false)

          isSectionStatePlantItem && (isSectionStatePlant = true)

          !isSectionStatePlantItem &&
            summary.push({
              center,
              tableCode,
              type,
              ...(isNursery && { nursery }),
              usedAcreage,
              emptyAcreage,
              acreage,
              year,
              occupancy,
            })

          isSectionStatePlantItem &&
            TableOverviewSelectVirtualItem.SectionStatePlant.map((SectionStatePlantItem: any) => {
              const woodSpecy = SectionStatePlantItem.Section.Origin?.WoodSpecy.name,
                naturalWoodArea = SectionStatePlantItem.Section.ApprovedArea?.NaturalWoodArea.name,
                forestVegetativeLevel = SectionStatePlantItem.Section.ApprovedArea?.ForestVegetativeLevel.name,
                age = SectionStatePlantItem.age
                  ? `${SectionStatePlantItem.age}`
                  : t('section.tabs.operation.tableState.ageNotSpecified'),
                amount = SectionStatePlantItem.amount
                  ? `${SectionStatePlantItem.amount} ks`
                  : t('section.tabs.operation.tableState.notSpecified'),
                areaAr = SectionStatePlantItem.areaAr ? `${SectionStatePlantItem.areaAr}` : 'Neobsazeno',
                inventory = SectionStatePlantItem.inventory
                  ? t('section.tabs.operation.tableState.yes')
                  : t('section.tabs.operation.tableState.no'),
                trimmed = SectionStatePlantItem.trimmed
                  ? t('section.tabs.operation.tableState.yes')
                  : t('section.tabs.operation.tableState.no'),
                section = `${SectionStatePlantItem.Section?.order} / ${SectionStatePlantItem.Section?.year}`

              let typeOperation = t('section.tabs.operation.tableState.readyForExport'),
                localization = t('section.tabs.operation.tableState.picked')

              if (SectionStatePlantItem.nursered) {
                typeOperation = t('section.tabs.operation.tableState.nursering')
              }
              if (SectionStatePlantItem.planted && !SectionStatePlantItem.nursered) {
                typeOperation = t('section.tabs.operation.tableState.plant')
              }

              if (SectionStatePlantItem.nursered) {
                localization = `${SectionStatePlantItem.NurseryOperation?.localization}`
              }
              if (SectionStatePlantItem.planted && !SectionStatePlantItem.nursered) {
                localization = `${SectionStatePlantItem.PlantOperation?.localization}`
              }

              summary.push({
                center,
                tableCode,
                type,
                ...(isNursery && { nursery }),
                usedAcreage,
                emptyAcreage,
                acreage,
                year,
                occupancy,
                woodSpecy,
                naturalWoodArea,
                forestVegetativeLevel,
                age,
                amount,
                typeOperation,
                localization,
                areaAr,
                inventory,
                trimmed,
                section,
              })

              updateColumnWidth(textToPixels(woodSpecy), summaryColumnsWidth, 'woodSpecy')
              updateColumnWidth(textToPixels(naturalWoodArea), summaryColumnsWidth, 'naturalWoodArea')
              updateColumnWidth(textToPixels(forestVegetativeLevel), summaryColumnsWidth, 'forestVegetativeLevel')
              updateColumnWidth(textToPixels(age), summaryColumnsWidth, 'age')
              updateColumnWidth(textToPixels(amount), summaryColumnsWidth, 'amount')
              updateColumnWidth(textToPixels(typeOperation), summaryColumnsWidth, 'typeOperation')
              updateColumnWidth(textToPixels(localization), summaryColumnsWidth, 'localization')
              updateColumnWidth(textToPixels(areaAr), summaryColumnsWidth, 'areaAr')
              updateColumnWidth(textToPixels(inventory), summaryColumnsWidth, 'inventory')
              updateColumnWidth(textToPixels(trimmed), summaryColumnsWidth, 'trimmed')
              updateColumnWidth(textToPixels(section), summaryColumnsWidth, 'section')
            })

          updateColumnWidth(textToPixels(center), summaryColumnsWidth, 'center')
          updateColumnWidth(textToPixels(tableCode), summaryColumnsWidth, 'tableCode')
          updateColumnWidth(textToPixels(type), summaryColumnsWidth, 'type')
          updateColumnWidth(textToPixels(nursery), summaryColumnsWidth, 'nursery')
          updateColumnWidth(textToPixels(usedAcreage), summaryColumnsWidth, 'usedAcreage')
          updateColumnWidth(textToPixels(emptyAcreage), summaryColumnsWidth, 'emptyAcreage')
          updateColumnWidth(textToPixels(acreage), summaryColumnsWidth, 'acreage')
          updateColumnWidth(textToPixels(year), summaryColumnsWidth, 'year')
          updateColumnWidth(textToPixels(occupancy), summaryColumnsWidth, 'occupancy')
        })
      })
    })

    const ws = XLSX.utils.json_to_sheet([])

    const headers = [
      [
        t('summary.headers.center'),
        customerModules.includes(EModule.Nursery)
          ? t('section.tabs.operation.tableState.table')
          : t('section.tabs.operation.stock'),
        t('section.tabs.operation.tableState.type'),
        t('section.tabs.operation.tableState.usedAreaAr'),
        t('section.tabs.operation.tableState.emptyAreaAr'),
        t('section.tabs.operation.tableState.acreage'),
        t('section.tabs.operation.tableState.year'),
        t('section.tabs.operation.tableState.occupancy'),
      ],
    ]

    isSectionStatePlant &&
      headers[0].push(
        t('section.tabs.operation.tableState.woodSpecies'),
        t('section.tabs.operation.tableState.naturalWoodArea'),
        t('section.tabs.operation.tableState.forestryVegetativeLevel'),
        t('section.tabs.operation.tableState.age'),
        t('section.tabs.operation.tableState.amount'),
        t('section.tabs.operation.tableState.type'),
        t('section.tabs.operation.tableState.localization'),
        t('section.tabs.operation.tableState.areaAr'),
        t('section.tabs.operation.tableState.inventory'),
        t('section.tabs.operation.tableState.trimmed'),
        t('section.tabs.operation.tableState.section')
      )

    // Add Nursery column if Nursery is in the list
    isNursery && headers[0].splice(3, 0, t('section.tabs.operation.tableState.nursery'))

    ws['!cols'] = [
      { wpx: summaryColumnsWidth.center },
      { wpx: summaryColumnsWidth.tableCode },
      { wpx: summaryColumnsWidth.type },
      { wpx: summaryColumnsWidth.usedAcreage },
      { wpx: summaryColumnsWidth.emptyAcreage },
      { wpx: summaryColumnsWidth.acreage },
      { wpx: summaryColumnsWidth.year },
      { wpx: summaryColumnsWidth.occupancy },
    ]

    isSectionStatePlant &&
      ws['!cols'].push(
        { wpx: summaryColumnsWidth.woodSpecy },
        { wpx: summaryColumnsWidth.naturalWoodArea },
        { wpx: summaryColumnsWidth.forestVegetativeLevel },
        { wpx: summaryColumnsWidth.age },
        { wpx: summaryColumnsWidth.amount },
        { wpx: summaryColumnsWidth.typeOperation },
        { wpx: summaryColumnsWidth.localization },
        { wpx: summaryColumnsWidth.areaAr },
        { wpx: summaryColumnsWidth.inventory },
        { wpx: summaryColumnsWidth.trimmed },
        { wpx: summaryColumnsWidth.section }
      )

    isNursery && ws['!cols'].splice(3, 0, { wpx: summaryColumnsWidth.nursery })

    XLSX.utils.sheet_add_aoa(ws, headers)
    XLSX.utils.sheet_add_json(ws, summary, { origin: 'A2', skipHeader: true })

    return ws
  }

  const wb = XLSX.utils.book_new()

  let createdWs: XLSX.WorkSheet = [],
    sheetTitle = ''

  switch (type) {
    case ESummaryTabType.Center:
      sheetTitle = t('summary.groupByCenters')
      createdWs = editTabsData(true, t('summary.headers.center'), t('summary.headers.woodSpecies'))
      break
    case ESummaryTabType.Wood:
      sheetTitle = t('summary.groupByWoodSpecies')
      createdWs = editTabsData(false, t('summary.headers.woodSpecies'), t('summary.headers.center'))
      break
    case ESummaryTabType.Table:
      sheetTitle = t('summary.groupByTables')
      createdWs = editTableTabData()
      break
  }

  XLSX.utils.book_append_sheet(wb, createdWs, sheetTitle)
  XLSX.writeFile(wb, `${t('summary.title')}.xlsx`)
}
