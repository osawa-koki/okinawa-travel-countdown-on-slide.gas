// @ts-nocheck

function main() {
  const OBJECT_KEY = 'OBJECT_ID'

  const properties = PropertiesService.getScriptProperties()
  const presentationId = properties.getProperty('PRESENTATION_ID')
  const slideId = properties.getProperty('SLIDE_ID')
  const objectId = properties.getProperty(OBJECT_KEY)

  if (presentationId == null || slideId == null) {
    Logger.log('Presentation ID or Slide ID is not set.')
    return
  }

  const presentation = SlidesApp.openById(presentationId)
  const slide = presentation.getSlides().find(slide => slide.getObjectId() === slideId)

  let shape = objectId ? slide.getShapes().find(shape => shape.getObjectId() === objectId) : null

  const { x, y, width, height } = { x: 20, y: 20, width: 400, height: 30 }

  const currentDateTime = new Date()
  const travelStartDateTime = new Date(properties.getProperty('TRAVEL_START_DATE'))

  const diff = travelStartDateTime - currentDateTime

  const text = (() => {
    if (diff < 0) null
    const days = Math.floor(diff / (1000 * 60 * 60 * 24))
    const hours = Math.floor((diff % (1000 * 60 * 60 * 24)) / (1000 * 60 * 60))
    const minutes = Math.floor((diff % (1000 * 60 * 60)) / (1000 * 60))
    return `Travel starts in ${days} days, ${hours} hours, ${minutes} minutes.`
  })()

  if (text == null) {
    if (shape == null) return
    shape.remove()
    properties.deleteProperty(OBJECT_KEY)
    Logger.log('Text box with ID: ' + objectId + ' has been removed.')
    return
  }

  if (shape != null) {
    shape.getText().setText(text)
    shape.setLeft(x)
    shape.setTop(y)
    shape.setWidth(width)
    shape.setHeight(height)
    Logger.log('Text box with ID: ' + objectId + ' has been updated.')
  } else {
    shape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, x, y, width, height)
    shape.getText().setText(text)
    const newObjectId = shape.getObjectId()
    properties.setProperty(OBJECT_KEY, newObjectId)
    Logger.log('New text box created with ID: ' + newObjectId)
  }
}
