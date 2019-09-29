require('../../css/easy_admin/calculation.scss')

if (null !== document.querySelector('.easyadmin.new-calculation, .easyadmin.edit-calculation')) {
    const calculator = document.querySelector('[name="calculation[calculator]"]')
    const update = () => {
        const calculatorClass = $(calculator).val()
        $('[data-calculator]').each((index, el) => {
            const enable = $(el).data('calculator') === calculatorClass
            $(el).toggle(enable)
            $(el).prop('disabled', !enable)
        })
    }

    $(calculator).on('change', update)
    update()
}
