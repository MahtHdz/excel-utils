export const routeNotFound = (req, res) => {
    res.status(404).json({
        message : 'Route not found'
    })
}

const apiController = {
    routeNotFound
}

export default apiController